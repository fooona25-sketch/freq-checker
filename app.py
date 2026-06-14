"""
Thai Satellite Frequency Cross-Check — web app (Phase 1)
ตรวจสอบความถี่ข่ายงานดาวเทียมต่างชาติ (API/A) เทียบกับข่ายไทย เพื่อสนับสนุนการทักท้วงตาม RR No. 9.3

Phase 1 = การเทียบความถี่ล้วน (deterministic) ไม่ใช้ Claude API
Phase 2 (ภายหลัง) = เพิ่มปุ่มร่างบันทึก/หนังสือ
"""
import tempfile
import pandas as pd
import streamlit as st
import engine

st.set_page_config(page_title="ตรวจสอบความถี่ดาวเทียม (ทักท้วง 9.3)", layout="wide")

def secret(key: str, default: str = "") -> str:
    try:
        return st.secrets.get(key, default)
    except Exception:
        return default

# ---------- ด่านรหัสผ่านอย่างง่าย (กันคนนอก) ----------
def check_password() -> bool:
    correct = secret("password", "changeme")
    if st.session_state.get("auth_ok"):
        return True
    pw = st.text_input("รหัสผ่าน", type="password")
    if pw == "":
        st.stop()
    if pw == correct:
        st.session_state["auth_ok"] = True
        return True
    st.error("รหัสผ่านไม่ถูกต้อง")
    st.stop()

check_password()

st.title("📡 ตรวจสอบความถี่ดาวเทียมเพื่อการทักท้วง (RR No. 9.3)")
st.caption("เปรียบเทียบความถี่ข่ายต่างชาติ (API/A จาก BR IFIC) กับฐานความถี่ข่ายงานดาวเทียมไทย")

# ---------- ฐานความถี่ไทย ----------
with st.sidebar:
    st.header("ฐานความถี่ข่ายไทย")
    src = st.radio("แหล่งข้อมูล", ["Google Sheet (แนะนำ)", "อัปโหลดไฟล์ Excel/CSV"])
    thai_df = None
    if src == "Google Sheet (แนะนำ)":
        default_url = secret("thai_sheet_csv_url", "")
        url = st.text_input("ลิงก์ Google Sheet (เผยแพร่เป็น CSV)", value=default_url,
                            help="ใน Google Sheet: File > Share > Publish to web > เลือก CSV แล้ววางลิงก์ที่ได้ที่นี่")
        if url:
            try:
                thai_df = engine.normalise_thai(pd.read_csv(url))
                st.success(f"โหลดฐานไทยแล้ว {len(thai_df)} แถบความถี่")
            except Exception as e:
                st.error(f"อ่าน Sheet ไม่สำเร็จ: {e}")
    else:
        up = st.file_uploader("ไฟล์ฐานความถี่ไทย", type=["xlsx", "csv"])
        if up:
            raw = pd.read_csv(up) if up.name.endswith(".csv") else pd.read_excel(up)
            thai_df = engine.normalise_thai(raw)
            st.success(f"โหลดฐานไทยแล้ว {len(thai_df)} แถบความถี่")

    st.divider()
    st.header("ตัวเลือกการเทียบ")
    section = st.selectbox("Special section ที่จะตรวจ", ["API/A", "CR/C", "Notif"], index=0)
    min_khz = st.number_input("ตัดทิ้งจุดทับซ้อนที่แคบกว่า (kHz)", min_value=0.0, value=0.0, step=10.0)
    only_downlink = st.checkbox("แสดงเฉพาะขาลง (Emission) ที่กระทบภาคพื้น", value=False)

# ---------- ข้อมูลต่างชาติ (IFIC) ----------
st.subheader("1) อัปโหลดฐานข้อมูล BR IFIC")
mdb = st.file_uploader("ไฟล์ ificXXXX.mdb ของรอบที่ต้องการตรวจ", type=["mdb"])

run = st.button("▶️ เริ่มเปรียบเทียบ", type="primary", disabled=not (mdb and thai_df is not None))
if not mdb:
    st.info("กรุณาอัปโหลดไฟล์ .mdb และตั้งค่าฐานความถี่ไทยทางแถบซ้ายก่อน")

if run:
    with st.spinner("กำลังอ่านฐานข้อมูลและเปรียบเทียบ..."):
        with tempfile.NamedTemporaryFile(suffix=".mdb", delete=False) as tf:
            tf.write(mdb.getbuffer())
            path = tf.name
        fgn = engine.extract_foreign(path, section)
        matches = engine.compute_overlaps(fgn, thai_df, min_overlap_khz=min_khz)
        if only_downlink:
            matches = matches[matches["ขาลง_กระทบภาคพื้น"] == "✔"].reset_index(drop=True)
        summary = engine.summarise(matches)

    c1, c2, c3 = st.columns(3)
    c1.metric("ข่ายต่างชาติในรอบนี้", fgn["satellite"].nunique())
    c2.metric("ข่ายที่ทับซ้อนกับไทย", 0 if matches.empty else summary["foreign_satellite"].nunique())
    c3.metric("จำนวนจุดทับซ้อน", len(matches))

    if matches.empty:
        st.success("ไม่พบความถี่ทับซ้อนตามเกณฑ์ที่ตั้งไว้")
    else:
        st.subheader("2) สรุป — ข่ายต่างชาติที่กระทบข่ายไทย")
        st.dataframe(summary, use_container_width=True, hide_index=True)
        st.subheader("3) รายละเอียดทุกจุดทับซ้อน")
        st.dataframe(matches, use_container_width=True, hide_index=True)
        st.download_button(
            "⬇️ ดาวน์โหลดผลเป็น Excel",
            engine.to_excel_bytes(summary, matches),
            file_name=f"overlap_{mdb.name.replace('.mdb','')}_{section.replace('/','-')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.caption("หมายเหตุ: ผลนี้คือการคัดกรองแถบความถี่ทับซ้อน (deterministic) สำหรับช่วยหาข่ายที่ต้องพิจารณา "
               "ไม่ใช่คำวินิจฉัยการรบกวนทางเทคนิคขั้นสุดท้าย")
