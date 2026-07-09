# -*- coding: utf-8 -*-
"""
ThaiSat Freq Checker — โหมด 9.3 (API/A)
ขอบเขต = portal CSV (รายการตีพิมพ์ทางการ ITU)  |  ความถี่ = ific.mdb
จับคู่ด้วย Notice ID (+ Targeted notice ID สำหรับ MOD), กรอง wic_no = เลข IFIC ของไฟล์
ฐานความถี่ไทย: Google Sheet (thai_sheet_csv_url ใน Secrets) หรืออัปโหลด Excel/CSV
"""
import streamlit as st
import subprocess, csv, io, tempfile, os
import pandas as pd
from collections import defaultdict, Counter
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(page_title="ThaiSat Freq Checker — 9.3", layout="wide")

# ---------------- secrets (safe) ----------------
def get_secret(key, default=None):
    """อ่าน secret อย่างปลอดภัย — ถ้าไม่มีไฟล์ secrets เลยก็ไม่ทำให้แอปพัง"""
    try:
        return st.secrets.get(key, default)
    except Exception:
        return default

# ---------------- password gate (robust) ----------------
def check_password():
    pw_secret = get_secret("password")
    # ถ้ายังไม่ได้ตั้งรหัสใน Secrets → เตือนแต่ให้เข้าได้ (กันจอขาว)
    if not pw_secret:
        st.warning('ยังไม่ได้ตั้งรหัสผ่านใน Secrets — เพิ่มบรรทัด  password = "รหัสของคุณ"  '
                   'ในหน้า Settings → Secrets เพื่อเปิดด่านรหัสผ่าน (ตอนนี้เข้าใช้ได้โดยไม่ต้องใส่รหัส)')
        return True
    if st.session_state.get("auth_ok"):
        return True
    pw = st.text_input("🔒 รหัสผ่าน", type="password")
    if pw:
        if pw == pw_secret:
            st.session_state["auth_ok"] = True
            return True
        st.error("รหัสผ่านไม่ถูกต้อง")
    return False

# ---------------- mdb helpers ----------------
def mdb_table(mdb_path, table):
    out = subprocess.run(["mdb-export", mdb_path, table],
                         capture_output=True, text=True).stdout
    return list(csv.DictReader(io.StringIO(out)))

def detect_ific(notice_rows):
    """เลข IFIC = ค่า wic_no ที่พบบ่อยสุดในตาราง notice"""
    vals = [n.get("wic_no", "") for n in notice_rows if n.get("wic_no")]
    return Counter(vals).most_common(1)[0][0] if vals else None

# ---------------- portal CSV ----------------
def read_portal_csv(uploaded):
    raw = uploaded.getvalue().decode("utf-8-sig", errors="replace").splitlines()
    lines = [l for l in raw if not l.lower().startswith("sep=")]
    delim = ";" if (lines and lines[0].count(";") >= lines[0].count(",")) else ","
    return list(csv.DictReader(lines, delimiter=delim))

def col(row, *names):
    for n in names:
        if n in row:
            return row[n]
    return ""

# ---------------- Thai base ----------------
def normalise_thai(df):
    df.columns = [c.strip().lower() for c in df.columns]
    need = {"network_name", "freq_start_mhz", "freq_end_mhz"}
    if not need.issubset(df.columns):
        raise ValueError("ฐานไทยต้องมีคอลัมน์: network_name, freq_start_mhz, freq_end_mhz")
    df = df.dropna(subset=["freq_start_mhz", "freq_end_mhz"]).copy()
    df["freq_start_mhz"] = pd.to_numeric(df["freq_start_mhz"], errors="coerce")
    df["freq_end_mhz"] = pd.to_numeric(df["freq_end_mhz"], errors="coerce")
    return df.dropna(subset=["freq_start_mhz", "freq_end_mhz"])

def load_thai(upload, sheet_url):
    """คืน (df, แหล่งที่มา) — ให้ความสำคัญกับไฟล์ที่อัปโหลดก่อน, ไม่งั้นใช้ Google Sheet"""
    if upload is not None:
        df = (pd.read_csv(upload) if upload.name.lower().endswith(".csv")
              else pd.read_excel(upload))
        return normalise_thai(df), "ไฟล์ที่อัปโหลด"
    if sheet_url:
        df = pd.read_csv(sheet_url)
        return normalise_thai(df), "Google Sheet"
    return None, None

# ---------------- core ----------------
def build_foreign_apia(mdb_path, portal_rows, ific):
    grp = mdb_table(mdb_path, "grp")
    gbn = defaultdict(list)
    for g in grp:
        if g.get("wic_no") == ific:            # group-level wic_no filter
            gbn[g["ntc_id"]].append(g)
    apia = [r for r in portal_rows if col(r, "Special section").strip() == "API/A"]
    foreign, nodata = [], []
    for r in apia:
        nid = col(r, "Notice ID")
        tgt = col(r, "Targeted notice ID")
        mid = next((x for x in [nid, tgt] if x and x in gbn), None)
        sat = col(r, "Satellite name")
        adm = col(r, "Notifying administration")
        if not mid:
            nodata.append((sat, adm, nid))
            continue
        for g in gbn[mid]:
            if g.get("freq_min"):
                foreign.append({
                    "sat": sat, "adm": adm, "notice_id": mid,
                    "beam": g.get("beam_name", ""), "emi": g.get("emi_rcp", ""),
                    "f_min": float(g["freq_min"]), "f_max": float(g["freq_max"]),
                })
    return pd.DataFrame(foreign), apia, nodata

def overlap(foreign, thai, min_khz=0.0):
    # min_khz = 0  -> รวม "แตะขอบ" (ทับกันพอดี 0 MHz) ด้วย
    # min_khz > 0  -> เอาเฉพาะทับจริงที่กว้างกว่าค่าที่ตั้ง (ตัดแตะขอบทิ้ง)
    min_mhz = min_khz / 1000.0
    rows = []
    for f in foreign.itertuples(index=False):
        for t in thai.itertuples(index=False):
            lo = max(f.f_min, t.freq_start_mhz)
            hi = min(f.f_max, t.freq_end_mhz)
            width = hi - lo
            if width >= min_mhz and width >= 0:
                rows.append({
                    "ดาวเทียมต่างชาติ": f.sat, "adm": f.adm,
                    "ทิศทาง": "ขาลง (E)" if f.emi == "E" else "ขาขึ้น (R)",
                    "ขาลงกระทบภาคพื้น": "✔" if f.emi == "E" else "",
                    "ประเภท": "ทับจริง" if width > 0 else "แตะขอบ (0 MHz)",
                    "beam": f.beam, "ตปท_f_min": round(f.f_min, 4), "ตปท_f_max": round(f.f_max, 4),
                    "ข่ายไทย": t.network_name,
                    "overlap_min": round(lo, 4), "overlap_max": round(hi, 4),
                    "overlap_MHz": round(width, 4),
                })
    return pd.DataFrame(rows)

def to_excel_bytes(summary, detail, nodata, ific):
    F = "Tahoma"
    hf = Font(name=F, bold=True, color="FFFFFF")
    hfill = PatternFill("solid", fgColor="1F3864")
    bd = Border(*[Side(style="thin", color="D0D0D0")] * 4)
    wb = Workbook()

    def style(ws):
        for c in ws[1]:
            c.font = hf; c.fill = hfill; c.border = bd
            c.alignment = Alignment(horizontal="center", vertical="center")
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.font = Font(name=F); c.border = bd
        ws.freeze_panes = "A2"
        if ws.max_row >= 1:
            ws.auto_filter.ref = ws.dimensions

    ws = wb.active; ws.title = f"สรุป (IFIC {ific})"
    ws.append(list(summary.columns))
    for _, r in summary.iterrows():
        ws.append(list(r))
    for col_, w in zip("ABCDEFG", [22, 8, 16, 10, 10, 10, 70]):
        ws.column_dimensions[col_].width = w
    style(ws)

    ws2 = wb.create_sheet("รายละเอียด")
    ws2.append(list(detail.columns))
    for _, r in detail.iterrows():
        ws2.append(list(r))
    for col_, w in zip("ABCDEFGHIJKL", [20, 8, 12, 14, 14, 16, 12, 12, 22, 12, 12, 12]):
        ws2.column_dimensions[col_].width = w
    style(ws2)

    if nodata:
        ws3 = wb.create_sheet("ไม่มี technical data")
        ws3.append(["ดาวเทียม", "adm", "Notice ID"])
        for s, a, n in nodata:
            ws3.append([s, a, n])
        for col_, w in zip("ABC", [24, 10, 16]):
            ws3.column_dimensions[col_].width = w
        style(ws3)

    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()

# ================= UI =================
st.title("🛰️ ตรวจความถี่ดาวเทียม — โหมด 9.3 (API/A)")

if not check_password():
    st.stop()

st.markdown(
    "อัปโหลด **2 ไฟล์**: ฐานข้อมูล `ific.mdb` (ความถี่) และ **portal CSV** "
    "(รายการตีพิมพ์ทางการ ITU = ขอบเขต) ระบบจะจับคู่ด้วย Notice ID"
)
c1, c2 = st.columns(2)
mdb_file = c1.file_uploader("1) ไฟล์ ific.mdb", type=["mdb"])
portal_file = c2.file_uploader("2) portal CSV (รายการตีพิมพ์)", type=["csv"])

# --- ฐานความถี่ไทย: Google Sheet หรืออัปโหลด ---
sheet_url = get_secret("thai_sheet_csv_url")
if sheet_url:
    st.success("ฐานความถี่ไทย: ใช้จาก **Google Sheet** ที่ตั้งไว้ใน Secrets "
               "(อัปโหลดไฟล์ด้านล่างเพื่อใช้แทนเฉพาะครั้งนี้ได้)")
else:
    st.info("ฐานความถี่ไทย: ยังไม่ได้ตั้ง Google Sheet — อัปโหลดไฟล์ด้านล่าง "
            'หรือเพิ่ม  thai_sheet_csv_url = "ลิงก์ CSV"  ใน Secrets')
thai_file = st.file_uploader("3) ฐานความถี่ไทย (Excel/CSV) — ข้ามได้ถ้าใช้ Google Sheet", type=["xlsx", "csv"])

with st.expander("ℹ️ วิธีตั้ง Google Sheet เป็นฐานความถี่ไทย"):
    st.markdown(
        "1. ใน Google ชีต: **ไฟล์ → แชร์ → เผยแพร่ไปยังเว็บ**\n"
        "2. เลือกรูปแบบ **CSV** → เผยแพร่ → คัดลอกลิงก์ (ลงท้าย `output=csv`)\n"
        "3. ใส่ใน **Settings → Secrets**:  `thai_sheet_csv_url = \"ลิงก์ที่คัดลอก\"`\n"
        "4. ชีตต้องมีคอลัมน์: **network_name, freq_start_mhz, freq_end_mhz**"
    )

min_khz = st.number_input("กรองจุดทับซ้อนที่เล็กกว่า (kHz) — 0 = เอาทุกจุด",
                          min_value=0.0, value=0.0, step=1.0)

if st.button("▶ ประมวลผล", type="primary"):
    if not mdb_file or not portal_file:
        st.error("ต้องอัปโหลดทั้ง ific.mdb และ portal CSV")
        st.stop()
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".mdb") as tf:
            tf.write(mdb_file.getvalue()); mdb_path = tf.name

        notice = mdb_table(mdb_path, "notice")
        ific = detect_ific(notice)
        portal_rows = read_portal_csv(portal_file)

        try:
            thai, thai_src = load_thai(thai_file, sheet_url)
        except Exception as e:
            st.error(f"อ่านฐานความถี่ไทยไม่สำเร็จ: {e}")
            st.stop()
        if thai is None:
            st.error("ยังไม่มีฐานความถี่ไทย — อัปโหลดไฟล์ หรือตั้งค่า Google Sheet ใน Secrets")
            st.stop()

        foreign, apia, nodata = build_foreign_apia(mdb_path, portal_rows, ific)
        os.unlink(mdb_path)

        st.caption(f"ฐานความถี่ไทย: {thai_src} ({len(thai)} ย่าน)")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("IFIC", ific)
        m2.metric("API/A (ขอบเขต)", len(apia))
        m3.metric("มีความถี่", foreign["sat"].nunique() if len(foreign) else 0)
        m4.metric("ไม่มี tech data", len(nodata))

        if nodata:
            st.warning("ข่ายที่ไม่มี technical data ในไฟล์รอบนี้ (publication-only): "
                       + ", ".join(s for s, _, _ in nodata))

        if len(foreign) == 0:
            st.info("ไม่พบความถี่สำหรับ API/A รอบนี้")
            st.stop()

        detail = overlap(foreign, thai, min_khz)
        if len(detail) == 0:
            st.success("ไม่มีความถี่ใดทับซ้อนกับฐานไทย")
            st.stop()

        n_real = int((detail["ประเภท"] == "ทับจริง").sum())
        n_edge = int((detail["ประเภท"] == "แตะขอบ (0 MHz)").sum())
        st.caption(f"จุดทับซ้อนทั้งหมด {len(detail)}  —  ทับจริง {n_real} | แตะขอบ (0 MHz) {n_edge}")

        summ = (detail.groupby(["ดาวเทียมต่างชาติ", "adm"])
                .agg(ข่ายไทยที่กระทบ=("ข่ายไทย", "nunique"),
                     ทับจริง=("ประเภท", lambda s: int((s == "ทับจริง").sum())),
                     แตะขอบ=("ประเภท", lambda s: int((s == "แตะขอบ (0 MHz)").sum())),
                     มีขาลง=("ขาลงกระทบภาคพื้น", lambda s: "✔" if (s == "✔").any() else ""),
                     รายชื่อข่ายไทย=("ข่ายไทย", lambda s: ", ".join(sorted(set(s)))))
                .reset_index())

        st.subheader("สรุปต่อข่าย")
        st.dataframe(summ, width="stretch")
        st.subheader("รายละเอียดจุดทับซ้อน")
        st.dataframe(detail, width="stretch")

        xls = to_excel_bytes(summ, detail, nodata, ific)
        st.download_button("⬇ ดาวน์โหลดผลเป็น Excel", xls,
                           file_name=f"IFIC{ific}_9.3_overlap.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาด: {e}")
