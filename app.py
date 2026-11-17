# -*- coding: utf-8 -*-
"""
ThaiSat Freq Checker
โหมด 1: 9.3 (API/A)  — ขอบเขตจาก portal CSV + ความถี่จาก ific.mdb + เทียบฐานไทย
โหมด 2: CR/C (9.7 / 9.12A / 9.21 / 9.14) — tr_aff_ntw (ชื่อข่ายไทยจาก BR) + provn/grp (ย่านความถี่)
ทุกโหมดกรอง wic_no = เลข IFIC ของไฟล์ (อ่านอัตโนมัติ)
ผลลัพธ์มีชีต "สำหรับบันทึก" จัดรูปแบบตามตารางบันทึกเสนอผู้บริหาร (1 แถว/ข่ายต่างชาติ,
ชื่อข่ายไทยเรียงลงบรรทัดในเซลล์เดียว พร้อมก๊อปลง Word)
"""
import streamlit as st
import subprocess, csv, io, tempfile, os, math, re
import pandas as pd
from collections import defaultdict, Counter
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---- ส่วนเสริม (กราฟชนย่าน + ร่างบันทึก Word) — ถ้าไลบรารียังไม่ติดตั้ง แอปหลักต้องทำงานต่อได้ ----
try:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib import font_manager as _fm
    from matplotlib.patches import Rectangle as _Rect
    _HAS_MPL = True
    for _f in ("fonts/Sarabun-Regular.ttf", "fonts/Sarabun-Bold.ttf",
               "Sarabun-Regular.ttf", "Sarabun-Bold.ttf"):
        if os.path.exists(_f):
            _fm.fontManager.addfont(_f)
    _THAI_FONT_OK = any(f.name == "Sarabun" for f in _fm.fontManager.ttflist)
    if _THAI_FONT_OK:
        plt.rcParams["font.family"] = "Sarabun"
except Exception:
    _HAS_MPL = False
    _THAI_FONT_OK = False
try:
    from docx import Document as _Docx
    from docx.shared import Pt as _Pt, Cm as _Cm, Emu as _Emu
    from docx.enum.text import WD_ALIGN_PARAGRAPH as _AL
    from docx.enum.table import WD_TABLE_ALIGNMENT as _TAL
    from docx.oxml.ns import qn as _qn
    from docx.oxml import OxmlElement as _OE
    _HAS_DOCX = True
except Exception:
    _HAS_DOCX = False

st.set_page_config(page_title="ThaiSat Freq Checker", layout="wide")
# ---------------- ธีม กสทช. ----------------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;600;700&display=swap');
html, body, [class*="st-"], button, input, label { font-family:'Sarabun',sans-serif !important; }
span[data-testid="stIconMaterial"], [class*="material-symbols"] { font-family:'Material Symbols Rounded' !important; }
header[data-testid="stHeader"] { background:#0E2A52; }
.nbtc-band { background:#0E2A52; color:#fff; padding:14px 22px;
  border-radius:8px; border-bottom:4px solid #C8102E; margin-bottom:10px; }
.nbtc-band h1 { font-size:1.45rem; margin:0; color:#fff; }
.nbtc-band p  { margin:3px 0 0; color:#CBD5E1; font-size:.9rem; }
.stButton>button[kind="primary"], .stDownloadButton>button { background:#0E2A52; border-color:#0E2A52; color:#fff; }
.stButton>button[kind="primary"]:hover, .stDownloadButton>button:hover { background:#C8102E; border-color:#C8102E; }
</style>
""", unsafe_allow_html=True)

if os.path.exists("nbtc_logo.png"):
    st.logo("nbtc_logo.png")
st.markdown('<div class="nbtc-band"><h1>ระบบตรวจสอบข่ายงานดาวเทียม BR IFIC</h1>'
            '<p>ส่วนประสานงานกิจการดาวเทียม · สำนักกิจการดาวเทียมสื่อสาร · สำนักงาน กสทช.</p></div>',
            unsafe_allow_html=True)

# ---------------- secrets (safe) ----------------
def get_secret(key, default=None):
    try:
        return st.secrets.get(key, default)
    except Exception:
        return default

# ---------------- password gate (robust) ----------------
def check_password():
    pw_secret = get_secret("password")
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

def portal_section_str(r):
    """API/A/14341 หรือ API/A/14116 MOD-1 จากคอลัมน์ portal"""
    s = f"{col(r,'Special section')}/{col(r,'Special section number')}"
    if col(r, "Type of revision") == "M":
        s += f" MOD-{col(r,'Revision number')}"
    return s

# ---------------- formatting ----------------
def fmt_lon(x):
    x = (str(x) if x is not None else "").strip()
    if not x:
        return "NGSO"
    try:
        v = round(float(x), 1)
    except ValueError:
        return x
    hemi = "E" if v >= 0 else "W"
    v = abs(v)
    return f"{int(v)} {hemi}" if v == int(v) else f"{v} {hemi}"

def thai_label(sat_name, long_nom):
    """ชื่อข่ายไทยสำหรับบันทึก: GSO -> 'ชื่อ (xx E)', NGSO -> 'ชื่อ (NGSO)'"""
    lon = (long_nom or "").strip()
    return f"{sat_name} ({fmt_lon(lon)})" if lon else f"{sat_name} (NGSO)"

def ssn_map(pub_rows, prefix):
    """ntc_id -> 'CR/C/6245 MOD-2' (เลือกเฉพาะ special section ตาม prefix)"""
    m = {}
    for r in pub_rows:
        if r.get("ssn_ref", "").startswith(prefix):
            s = f"{prefix}/{r['ssn_no']}"
            if r.get("ssn_rev") == "M":
                s += f" MOD-{r.get('ssn_rev_no','')}"
            m[r["ntc_id"]] = s
    return m

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
    if upload is not None:
        df = (pd.read_csv(upload) if upload.name.lower().endswith(".csv")
              else pd.read_excel(upload))
        return normalise_thai(df), "ไฟล์ที่อัปโหลด"
    if sheet_url:
        return normalise_thai(pd.read_csv(sheet_url)), "Google Sheet"
    return None, None

# ---------------- ตาราง NFAT (บริการภาคพื้นของไทย) ----------------
_NFAT_UNIT = {"khz": 1e-3, "mhz": 1.0, "ghz": 1e3}
# ไม่ใช่บริการภาคพื้นตามนิยาม RR (No. 1.6): กิจการดาวเทียมทุกชนิด, space research/operation,
# radio astronomy — และตัดแถว Not Allocation
_NFAT_EXCL = re.compile(r"satellite|space research|space operation|radio astronomy|"
                        r"inter-satellite|not allocation", re.I)

def normalise_nfat(df):
    df.columns = [str(c).strip() for c in df.columns]
    need = {"Start_Frequency", "Stop_Frequency", "FrequencyUnit", "EngService", "AllocationType"}
    if not need.issubset(df.columns):
        raise ValueError("ตาราง NFAT ต้องมีคอลัมน์: Start_Frequency, Stop_Frequency, "
                         "FrequencyUnit, EngService, AllocationType")
    mul = df["FrequencyUnit"].astype(str).str.strip().str.lower().map(_NFAT_UNIT)
    df = df.assign(
        f1_mhz=pd.to_numeric(df["Start_Frequency"], errors="coerce") * mul,
        f2_mhz=pd.to_numeric(df["Stop_Frequency"], errors="coerce") * mul,
    ).dropna(subset=["f1_mhz", "f2_mhz"])
    return df

def load_nfat(upload, sheet_url):
    if upload is not None:
        df = (pd.read_csv(upload) if upload.name.lower().endswith(".csv")
              else pd.read_excel(upload))
        return normalise_nfat(df), "ไฟล์ที่อัปโหลด"
    if sheet_url:
        return normalise_nfat(pd.read_csv(sheet_url)), "Google Sheet"
    return None, None

def nfat_services(nfat, a, b):
    """บริการภาคพื้นของไทยที่จัดสรรทับช่วง a–b MHz (ตาม NFAT): Primary = ตัวพิมพ์ใหญ่,
    Secondary = ต่อท้าย (รอง) ตามธรรมเนียม RR"""
    if nfat is None:
        return "(ไม่ได้แนบตาราง NFAT)"
    m = nfat[(nfat["f1_mhz"] < b) & (nfat["f2_mhz"] > a)].sort_values("f1_mhz")
    out, seen = [], set()
    for _, r in m.iterrows():
        e = str(r["EngService"]).strip()
        if not e or _NFAT_EXCL.search(e):
            continue
        pri = str(r["AllocationType"]).strip().lower() == "primary"
        k = (e.lower(), pri)
        if k in seen:
            continue
        seen.add(k)
        out.append(e.upper() if pri else f"{e} (รอง)")
    return ", ".join(out) if out else "ไม่มีบริการภาคพื้นจัดสรรในย่านนี้ (ตาม NFAT)"

# ---------------- Excel writer (ทั่วไป) ----------------
def build_excel(sheets, green_sheets=()):
    """sheets = list ของ (ชื่อชีต, DataFrame, [ความกว้างคอลัมน์]);
    green_sheets = ชีตที่ให้ระบายช่องค่า X (เขียว) / x (เขียวอ่อน) แบบชีตสรุป"""
    F = "Tahoma"
    hf = Font(name=F, bold=True, color="FFFFFF")
    hfill = PatternFill("solid", fgColor="1F3864")
    gfill = PatternFill("solid", fgColor="92D050")
    gfill2 = PatternFill("solid", fgColor="E2EFDA")
    bd = Border(*[Side(style="thin", color="D0D0D0")] * 4)
    wb = Workbook(); wb.remove(wb.active)
    for title, df, widths in sheets:
        ws = wb.create_sheet(title[:31])
        ws.append(list(df.columns))
        for _, r in df.iterrows():
            ws.append(list(r))
        for c in ws[1]:
            c.font = hf; c.fill = hfill; c.border = bd
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        mark_me = title in green_sheets
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.font = Font(name=F); c.border = bd
                c.alignment = Alignment(vertical="top", wrap_text=True)
                if mark_me and c.value in ("X", "x"):
                    c.fill = gfill if c.value == "X" else gfill2
                    c.alignment = Alignment(horizontal="center", vertical="center")
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"
        if len(df):
            ws.auto_filter.ref = ws.dimensions
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()

def thai_sort(names):
    """ลำดับรายชื่อข่ายไทยตามแบบบันทึก: NGSO ก่อน → GSO (THAICOM ฯลฯ) → THAISAT ปิดท้าย"""
    def k(n):
        u = str(n).upper()
        return (0 if "(NGSO" in u.replace(" ", "") else 1,
                1 if u.startswith("THAISAT") else 0, u)
    return sorted(set(names), key=k)

def nl(items):
    return "\n".join(items) if items else "—"

# ================= โหมด 9.3 (API/A) =================
def build_foreign_apia(mdb_path, portal_rows, ific):
    grp = mdb_table(mdb_path, "grp")
    gbn = defaultdict(list)
    for g in grp:
        if g.get("wic_no") == ific:
            gbn[g["ntc_id"]].append(g)
    apia = [r for r in portal_rows if col(r, "Special section").strip() == "API/A"]
    foreign, nodata = [], []
    for r in apia:
        nid = col(r, "Notice ID"); tgt = col(r, "Targeted notice ID")
        mid = next((x for x in [nid, tgt] if x and x in gbn), None)
        sat = col(r, "Satellite name"); adm = col(r, "Notifying administration")
        if not mid:
            nodata.append((sat, adm, nid, portal_section_str(r)))
            continue
        for g in gbn[mid]:
            if g.get("freq_min"):
                foreign.append({
                    "sat": sat, "adm": adm, "notice_id": mid,
                    "section": portal_section_str(r),
                    "orbit": col(r, "Orbital position") or "NGSO",
                    "beam": g.get("beam_name", ""), "emi": g.get("emi_rcp", ""),
                    "f_min": float(g["freq_min"]), "f_max": float(g["freq_max"]),
                })
    return pd.DataFrame(foreign), apia, nodata

def overlap(foreign, thai, min_khz=0.0):
    # min_khz = 0 -> รวม "แตะขอบ" (0 MHz) ด้วย / >0 -> เฉพาะทับจริงที่กว้างพอ
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
                    "beam": f.beam,
                    "ตปท_f_min": round(f.f_min, 4), "ตปท_f_max": round(f.f_max, 4),
                    "ข่ายไทย": t.network_name,
                    "overlap_min": round(lo, 4), "overlap_max": round(hi, 4),
                    "overlap_MHz": round(width, 4),
                })
    return pd.DataFrame(rows)

def memo_93(detail, foreign, nodata):
    """ชีต 'สำหรับบันทึก' — 1 แถว/ข่ายต่างชาติ ตามตารางบันทึก ข้อ 9.3"""
    meta = (foreign.groupby("sat")
            .agg(orbit=("orbit", "first"), section=("section", "first"),
                 adm=("adm", "first"),
                 has_dl=("emi", lambda s: "X" if (s == "E").any() else ""))
            .to_dict("index"))
    rows = []
    for sat, g in detail.groupby("ดาวเทียมต่างชาติ"):
        real = thai_sort(g.loc[g["ประเภท"] == "ทับจริง", "ข่ายไทย"])
        edge = thai_sort(set(g.loc[g["ประเภท"] == "แตะขอบ (0 MHz)", "ข่ายไทย"]) - set(real))
        m = meta.get(sat, {})
        rows.append({
            "ข่ายงานดาวเทียม": sat,
            "Nominal Longitude": m.get("orbit", "NGSO"),
            "Special section": m.get("section", ""),
            "ประเทศ": m.get("adm", ""),
            "ข่ายไทยที่กระทบ (S)": nl(real),
            "แตะขอบ (0 MHz)": nl(edge),
            "T (มีขาลง)": m.get("has_dl", ""),
            "มาตรา": "9.3",
        })
    for sat, adm, nid, sec in nodata:
        rows.append({"ข่ายงานดาวเทียม": sat, "Nominal Longitude": "NGSO",
                     "Special section": sec, "ประเทศ": adm,
                     "ข่ายไทยที่กระทบ (S)": "— ไม่มี technical data ในไฟล์รอบนี้ —",
                     "แตะขอบ (0 MHz)": "", "T (มีขาลง)": "", "มาตรา": "9.3"})
    rows.sort(key=lambda d: int(sec_num(d["Special section"]) or 999999))
    return pd.DataFrame(rows)

# ================= โหมด CR/C =================
def crc_prov_norm(p):
    if p.startswith("A30"):          # A30#7.1 / A30A#7.1(-CII) → AP30/30A
        return "AP30"
    return "9.7" if p.startswith("9.7") else p

def sec_num(sec):
    """'CR/C/6176 MOD-1' -> '6176'"""
    try:
        return sec.split("/")[-1].split()[0]
    except Exception:
        return ""

# ================= เครื่องคิด pfd สำหรับ 9.14 =================
# เกณฑ์: RR Appendix 5 Annex 1 — §1.1 (ต่ำกว่า 1 GHz) และ Table 5-2 (1–3 GHz, Rev.WRC-19)
# (f_lo, f_hi, ref_bw, P_gso, r_gso, P_ngso, r_ngso)  — r=None = ค่าคงที่ทุกมุม
PFD_TH = [
    (137.0,   138.0,  "4kHz", -125, None, -125, None),   # AP5 §1.1
    (400.15,  401.0,  "4kHz", -125, None, -125, None),   # AP5 §1.1
    (1518.0,  1525.0, "1MHz", -128, 0.5,  -128, 0.5),    # Table 5-2
    (1525.0,  1530.0, "1MHz", -128, 0.5,  -128, 0.5),    # Table 5-2
    (2160.0,  2200.0, "1MHz", -128, 0.5,  -123, 0.5),    # Table 5-2 (NGSO ผ่อนกว่า)
    (2483.5,  2500.0, "1MHz", -128, 0.5,  -126, 0.65),   # Table 5-2 (MSS; RDSS มีชุดค่าแยก)
]
RE_KM, RGSO_KM = 6371.0, 42164.0
# กรอบพิกัดไทย: จุดสุดเขตแดน 4 ทิศ (อ้างอิงข้อมูลกรมแผนที่ทหาร)
TH_LAT = (5.610, 20.468)
TH_LON = (97.350, 105.617)

def _spread_db(d_km):
    return 10 * math.log10(4 * math.pi * (d_km * 1000.0) ** 2)

def _slant_km(theta_deg, h_km):
    t = math.radians(theta_deg); r = (RE_KM + h_km) / RE_KM
    return RE_KM * (math.sqrt(r * r - math.cos(t) ** 2) - math.sin(t))

def _T(theta, P, r):
    if r is None: return P
    if theta <= 5: return P
    if theta <= 25: return P + r * (theta - 5)
    return P + 20 * r

_gso_cache = {}
def _gso_points(slon):
    """จุด (มุมเงย ε, ระยะ d, lat, lon) ทั่วกรอบพิกัดไทย สำหรับดาวเทียม GSO ที่ลองจิจูด slon"""
    if slon in _gso_cache: return _gso_cache[slon]
    pts = []
    la = TH_LAT[0]
    while la <= TH_LAT[1] + 1e-9:
        lo = TH_LON[0]
        while lo <= TH_LON[1] + 1e-9:
            cp = math.cos(math.radians(la)) * math.cos(math.radians(lo - slon))
            psi = math.acos(max(-1.0, min(1.0, cp)))
            if psi > 1e-6:
                e = math.degrees(math.atan((cp - RE_KM / RGSO_KM) / math.sin(psi)))
                if e > 0:
                    pts.append((e, math.sqrt(RE_KM**2 + RGSO_KM**2 - 2*RE_KM*RGSO_KM*cp), la, lo))
            lo += 0.25
        la += 0.25
    _gso_cache[slon] = pts
    return pts

def _th_segments(f1, f2, is_gso):
    """แบ่งย่านที่ตีพิมพ์เป็นส่วนที่มีเกณฑ์ (พร้อมค่า) และส่วนที่ไม่มีเกณฑ์"""
    segs, cur = [], f1
    for lo, hi, bw, Pg, rg, Pn, rn in PFD_TH:
        a, b = max(f1, lo), min(f2, hi)
        if b > a:
            segs.append((a, b, bw, (Pg if is_gso else Pn), (rg if is_gso else rn),
                         "AP5 §1.1" if bw == "4kHz" and hi <= 1000 else "Table 5-2"))
    segs.sort()
    un, cur = [], f1
    for a, b, *_ in segs:
        if a > cur: un.append((cur, a))
        cur = max(cur, b)
    if cur < f2: un.append((cur, f2))
    return segs, un

def pfd_worst(kind, par, E_bw, P, r):
    """หา margin แย่สุด: kind='NGSO' (par=perigee km, กวาด θ 0–90°) หรือ 'GSO' (par=ลองจิจูด, กวาดกรอบไทย)
    คืน (margin, θ*, d, จุด) หรือ None ถ้า GSO มองไม่เห็นจากไทย"""
    best = None
    if kind == "NGSO":
        t = 0.0
        while t <= 90.0 + 1e-9:
            d = _slant_km(t, par)
            m = E_bw - _spread_db(d) - _T(t, P, r)
            if best is None or m > best[0]: best = (m, t, d, None)
            t += 0.1
    else:
        for e, d, la, lo in _gso_points(par):
            m = E_bw - _spread_db(d) - _T(e, P, r)
            if best is None or m > best[0]: best = (m, e, d, (la, lo))
    return best

# ================= ชีตสรุปรวมทั้งฉบับ (ฟอร์มตารางทบทวนของ ปบ.) =================

SUM_PROVS = ["9.7", "9.12", "9.12A", "9.21/A", "9.21/B", "9.21/C", "9.14", "9.3", "AP30"]

def build_summary(portal_rows, memo93, R):
    """1 แถว/ข่ายที่ตีพิมพ์ (ทุกตัวใน portal) × ช่องมาตรา: X = เข้าเกณฑ์/กระทบไทย,
    x = แตะขอบ/เฉียด — Remark เว้นว่างให้กรอกเอง"""
    marks = defaultdict(dict)

    def put(sn, prov, val="X"):
        if not sn:
            return
        if marks[sn].get(prov) != "X":
            marks[sn][prov] = val

    for key, prov in (("m97", "9.7"), ("m912", "9.12"),
                      ("m912a", "9.12A"), ("ap30", "AP30")):
        df = R[key]
        for d in (df.to_dict("records") if len(df) else []):
            put(sec_num(d["Special section"]), prov)
    m921 = R["m921"]
    for d in (m921.to_dict("records") if len(m921) else []):
        prov, sn = d["มาตรา"], sec_num(d["Special section"])
        if prov in ("9.21/A", "9.21/B"):
            real = str(d.get("ข่ายไทยที่ทับ (ทับจริง)", "")).strip()
            edge = str(d.get("แตะขอบ (0 MHz)", "")).strip()
            if real not in ("", "—"):
                put(sn, prov, "X")
            elif edge not in ("", "—"):
                put(sn, prov, "x")
        else:
            put(sn, prov, "X")
    m914 = R["m914"]
    for d in (m914.to_dict("records") if len(m914) else []):
        put(sec_num(d["Special section"]), "9.14")
    if memo93 is not None and len(memo93):
        for d in memo93.to_dict("records"):
            sn = sec_num(d["Special section"])
            s = str(d.get("ข่ายไทยที่กระทบ (S)", "")).strip()
            e = str(d.get("แตะขอบ (0 MHz)", "")).strip()
            if s and not s.startswith("—"):
                put(sn, "9.3", "X")
            elif e and e != "—":
                put(sn, "9.3", "x")

    rows, seen = [], set()
    for r in portal_rows:
        sn = col(r, "Special section number").strip()
        if not sn or sn in seen:
            continue
        seen.add(sn)
        m = marks.get(sn, {})
        rows.append({
            "ADM": col(r, "Notifying administration"),
            "SATELLITE NETWORK": col(r, "Satellite name"),
            "SPECIAL SECTION": portal_section_str(r),
            "Longti": col(r, "Orbital position") or "NGSO",
            **{p: m.get(p, "") for p in SUM_PROVS},
            "Remark": "",
        })
    cols = ["ADM", "SATELLITE NETWORK", "SPECIAL SECTION", "Longti"] + SUM_PROVS + ["Remark"]
    df = pd.DataFrame(rows, columns=cols)
    if len(df):
        df = df.sort_values(["ADM", "SATELLITE NETWORK"]).reset_index(drop=True)
    return df

# ================= ผังการชนย่านความถี่รายข่าย =================

def collision_fig(name, meta, thai, ific):
    """วาดผังชนย่าน: แถวบน = ข่ายต่างชาติ, แถวล่าง = ข่ายไทยที่ทับ/แตะขอบ
    คืน (fig, None) หรือ (None, ข้อความเหตุผล)"""
    NAVY, CORAL, FOREIGN, RED = "#0E2A52", "#D85A30", "#7A1616", "#C62828"
    # รวมช่วงทิศเดียวกันที่ต่อเนื่อง
    sym = {"E": "↓", "R": "↑"}
    F = []
    for a, b, d in sorted(meta["bands"]):
        s = "↕" if len(d) > 1 else sym.get(d, "·")
        if F and a <= F[-1][1] + 1e-6 and F[-1][2] == s:
            F[-1][1] = max(F[-1][1], b)
        else:
            F.append([a, b, s])
    # ฐานไทย: ชนิด/ตำแหน่ง/ชื่อ
    tdf = thai.copy()
    tdf["kind"] = tdf["network_name"].astype(str).str.upper().str.contains(
        r"\(\s*NGSO\s*\)").map({True: "NGSO", False: "GSO"})
    tdf["slot"] = tdf["network_name"].astype(str).str.extract(r"\(([\d.]+E)\)")[0].fillna("")
    tdf["nm"] = tdf["network_name"].astype(str).str.replace(
        r"\s*\((?:NGSO|[\d.]+E)\)\s*$", "", regex=True).str.strip()
    scope = meta.get("scope", "all")
    if scope == "gso":
        tdf = tdf[tdf["kind"] == "GSO"]
    elif scope == "ngso":
        tdf = tdf[tdf["kind"] == "NGSO"]
    scope_txt = {"all": "เทียบฐานข่ายดาวเทียมไทยทั้งหมด",
                 "gso": "เทียบเฉพาะข่าย GSO ไทย (นิยาม 9.21/A)",
                 "ngso": "เทียบเฉพาะข่าย NGSO ไทย (นิยาม 9.21/B)"}[scope]
    hits = []
    for a, b, s in F:
        for t in tdf.itertuples(index=False):
            lo, hi = max(a, t.freq_start_mhz), min(b, t.freq_end_mhz)
            if hi > lo + 1e-9:
                hits.append((a, b, t.nm, t.kind, t.slot, lo, hi, "ทับ"))
            elif abs(hi - lo) <= 1e-9 and (abs(a - t.freq_end_mhz) < 1e-6 or abs(b - t.freq_start_mhz) < 1e-6):
                hits.append((a, b, t.nm, t.kind, t.slot, lo, lo, "แตะขอบ"))
    if not hits:
        return None, f"รายการนี้ไม่มีย่านที่ทับหรือแตะขอบ ({scope_txt})"

    inv = sorted({(h[0], h[1]) for h in hits})
    clus = [[list(inv[0])]]
    for a, b in inv[1:]:
        pb = max(x[1] for x in clus[-1])
        if (a - pb) <= max(30, pb * 0.12):
            clus[-1].append([a, b])
        else:
            clus.append([[a, b]])

    def _bn(f):
        for nm, x in [("VHF", 300), ("UHF", 1000), ("L-band", 1710), ("S-band", 4000),
                      ("C-band", 7100), ("X-band", 10700), ("Ku-band", 17300),
                      ("Ka-band", 33000)]:
            if f < x:
                return nm
        return "Q/V"

    panels = []
    for c in clus:
        lo = min(x[0] for x in c); hi = max(x[1] for x in c)
        pad = max((hi - lo) * 0.06, 0.6)
        w0, w1 = lo - pad, hi + pad
        th = sorted({(h[2], h[3], h[4]) for h in hits if lo <= h[0] and h[1] <= hi},
                    key=lambda t: (t[1] == "GSO", t[2], t[0]))
        panels.append((w0, w1, hi, th, [h for h in hits if lo <= h[0] and h[1] <= hi]))

    rowcount = sum(1 + len(p[3]) for p in panels)
    fig_h = 1.15 + 0.92 * len(panels) + 0.265 * rowcount
    fig = plt.figure(figsize=(9.5, fig_h))
    fig.text(0.5, 1 - 0.28 / fig_h,
             f"ผังการชนย่านความถี่ — {name} ({meta['adm']}, {meta['orbit']})",
             ha="center", va="top", fontsize=14, fontweight="bold", color=NAVY)
    fig.text(0.5, 1 - 0.56 / fig_h,
             f"BR IFIC {ific} · {scope_txt} · ย่านเฉพาะที่ตีพิมพ์ภายใต้มาตรานี้ · "
             "↓ ขาลง · ↑ ขาขึ้น · แถบแดงโปร่ง = ช่วงที่ทับ · ◆ = แตะขอบ (0 MHz)",
             ha="center", va="top", fontsize=7.8, color="#555")
    y_cursor = 1 - 0.95 / fig_h
    for w0, w1, hi, th, hh in panels:
        nrow = 1 + len(th)
        ph = (0.265 * nrow) / fig_h
        ax = fig.add_axes([0.21, y_cursor - ph, 0.76, ph])
        y_cursor -= ph + 0.82 / fig_h
        ax.set_xlim(w0, w1); ax.set_ylim(nrow - 0.5, -0.5)
        a0, a1 = _bn(w0), _bn(hi)
        ax.set_title(f"{a0 if a0 == a1 else a0 + ' – ' + a1} · {w0:,.6g} – {w1:,.6g} MHz",
                     loc="left", fontsize=9, fontweight="bold", color="#6B6355", pad=3)
        uni = []
        for ia, ib in sorted((h[5], h[6]) for h in hh if h[7] == "ทับ"):
            if uni and ia <= uni[-1][1] + 1e-9:
                uni[-1][1] = max(uni[-1][1], ib)
            else:
                uni.append([ia, ib])
        for k, (ia, ib) in enumerate(uni):
            ax.axvspan(ia, ib, color=RED, alpha=0.12, zorder=1)
            for x in (ia, ib):
                ax.axvline(x, color=RED, lw=0.7, ls=(0, (2, 2)), alpha=0.6, zorder=2)
            lab = f"ทับ {ia:,.6g}–{ib:,.6g} ({ib - ia:,.6g} MHz)"
            yoff = -0.30 if len(uni) > 2 and k % 2 else -0.17
            ax.text((ia + ib) / 2, yoff, lab, transform=ax.get_xaxis_transform(),
                    ha="center", va="top", fontsize=6.6, color=RED, fontweight="bold",
                    clip_on=False)
        edges = sorted({h[5] for h in hh if h[7] == "แตะขอบ"})
        if edges:
            ax.text(1.0, 1.03, "◆ แตะขอบ: " + " · ".join(f"{e:,.6g}" for e in edges) + " MHz",
                    transform=ax.transAxes, ha="right", va="bottom", fontsize=6.6,
                    color="#8F6400")
        for a, b, s in F:
            if b < w0 or a > w1:
                continue
            aa, bb = max(a, w0), min(b, w1)
            bb = max(bb, aa + (w1 - w0) * 0.004)
            ax.add_patch(_Rect((aa, -0.30), bb - aa, 0.60, color=FOREIGN, zorder=3))
            if bb - aa > (w1 - w0) * 0.055:
                ax.text((aa + bb) / 2, 0, s, ha="center", va="center", fontsize=7.5,
                        color="white", fontweight="bold", zorder=4)
        for i, (nm_, kd, sl) in enumerate(th, start=1):
            sub = tdf[(tdf.nm == nm_) & (tdf.kind == kd) & (tdf.slot == sl)]
            col_ = CORAL if kd == "NGSO" else NAVY
            for t in sub.itertuples(index=False):
                a, b = t.freq_start_mhz, t.freq_end_mhz
                if b < w0 or a > w1:
                    continue
                aa, bb = max(a, w0), min(b, w1)
                bb = max(bb, aa + (w1 - w0) * 0.004)
                ax.add_patch(_Rect((aa, i - 0.30), bb - aa, 0.60, color=col_, zorder=3))
            for h in hh:
                if h[2] == nm_ and h[7] == "แตะขอบ":
                    ax.plot(h[5], i, marker="D", ms=4.5, color="#E6A817", zorder=5,
                            markeredgecolor="#8F6400", markeredgewidth=0.4)
        labels = [f"{name}  ({meta['adm']})"] + [
            nm_ + ("  · " + sl.replace("E", "°E") if sl else "  · NGSO") for nm_, kd, sl in th]
        ax.set_yticks(range(nrow)); ax.set_yticklabels(labels, fontsize=7.6)
        kinds = ["F"] + [kd for _, kd, _ in th]
        for t, kd in zip(ax.get_yticklabels(), kinds):
            t.set_color(FOREIGN if kd == "F" else (CORAL if kd == "NGSO" else NAVY))
            if kd == "F":
                t.set_fontweight("bold")
        ax.tick_params(axis="y", length=0)
        ax.tick_params(axis="x", labelsize=6.6, length=2)
        ax.ticklabel_format(axis="x", style="plain", useOffset=False)
        ax.grid(axis="x", color="#E2DED4", lw=0.5)
        ax.set_axisbelow(True)
        for s_ in ["top", "right", "left"]:
            ax.spines[s_].set_visible(False)
        ax.spines["bottom"].set_color("#999")
    return fig, None

def charts_pdf(charts, thai, ific, progress=None):
    """PDF รวมเล่ม: หน้าแรกเป็นสารบัญครบทุกรายการที่ตรวจ (ไม่พบการทับ = ระบุชัด)
    ตามด้วยผังหน้าละรายการ เรียง 9.3 → 9.21/A → 9.21/B ตามเลข section"""
    from matplotlib.backends.backend_pdf import PdfPages
    import math as _m
    order = sorted(charts, key=lambda k: (charts[k].get("prov") != "9.3",
                                          charts[k].get("secn", 999999), k))
    t0 = thai.copy()
    t0["kind"] = t0["network_name"].astype(str).str.upper().str.contains(
        r"\(\s*NGSO\s*\)").map({True: "NGSO", False: "GSO"})
    res = []
    for k in order:
        m = charts[k]
        sc = m.get("scope", "all")
        td = t0 if sc == "all" else t0[t0["kind"] == ("GSO" if sc == "gso" else "NGSO")]
        nh = ne = 0
        for a, b, _d in m["bands"]:
            for t in td.itertuples(index=False):
                lo, hi = max(a, t.freq_start_mhz), min(b, t.freq_end_mhz)
                if hi > lo + 1e-9:
                    nh += 1
                elif abs(hi - lo) <= 1e-9 and (abs(a - t.freq_end_mhz) < 1e-6
                                               or abs(b - t.freq_start_mhz) < 1e-6):
                    ne += 1
        res.append((k, nh, ne))
    NAVY2, RED2, GRN2 = "#0E2A52", "#B3261E", "#2E7D32"
    GROUPS = [("9.3", "มาตรา 9.3 — API/A  (เทียบฐานข่ายดาวเทียมไทยทั้งหมด)"),
              ("9.21/A", "มาตรา 9.21/A  (เทียบเฉพาะข่าย GSO ไทย)"),
              ("9.21/B", "มาตรา 9.21/B  (เทียบเฉพาะข่าย NGSO ไทย)")]
    # จัดบรรทัด: หัวกลุ่ม + รายการ แล้วแบ่งหน้า
    lines = []
    for pv, gt in GROUPS:
        ent = [(k, h, e) for k, h, e in res if charts[k].get("prov") == pv]
        if not ent:
            continue
        lines.append(("g", f"{gt} — {len(ent)} รายการ"))
        lines.extend(("r", k, h, e) for k, h, e in ent)
    CAP = 30.0
    pages, cur, used = [], [], 0.0
    for ln in lines:
        w = 1.6 if ln[0] == "g" else 1.0
        if used + w > CAP and cur:
            pages.append(cur); cur, used = [], 0.0
        cur.append(ln); used += w
    if cur:
        pages.append(cur)
    n_idx = max(1, len(pages))
    ordered_rows = [ln[1:] for ln in lines if ln[0] == "r"]
    page_no, p = {}, n_idx + 1
    for k, h, e in ordered_rows:
        if h + e:
            page_no[k] = p; p += 1

    import datetime as _dt
    buf = io.BytesIO()
    pdf = PdfPages(buf)
    for ip, pls in enumerate(pages):
        fig = plt.figure(figsize=(9.5, 12.2))
        fig.patches.append(_Rect((0, 0.945), 1, 0.055, color=NAVY2,
                                 transform=fig.transFigure, zorder=1))
        fig.text(0.055, 0.9725, "สารบัญผังการชนย่านความถี่" + ("  (ต่อ)" if ip else ""),
                 fontsize=14, fontweight="bold", color="white", va="center", zorder=2)
        fig.text(0.945, 0.9725, f"BR IFIC {ific}  ·  จัดทำ {_dt.date.today():%d/%m/%Y}",
                 fontsize=8.5, color="#C9D4E6", ha="right", va="center", zorder=2)
        fig.text(0.055, 0.932, "ครอบคลุมทุกรายการที่ ปบ. เทียบความถี่เอง — รายการที่ไม่พบการทับ/แตะขอบ "
                 "แสดงผลการตรวจไว้โดยไม่มีหน้าแผนภาพ", fontsize=7.3, color="#777")
        XN, XA, XO, XR, XP = 0.055, 0.50, 0.575, 0.905, 0.945
        y = 0.905
        def colhead(yy):
            for x_, t_, ha_ in ((XN, "ข่ายงานดาวเทียม", "left"), (XA, "ADM", "left"),
                                (XO, "วงโคจร", "left"), (XR, "ผลการตรวจ", "right"),
                                (XP, "หน้า", "right")):
                fig.text(x_, yy, t_, fontsize=7.6, fontweight="bold", color="#6B6355", ha=ha_)
            fig.lines.append(plt.Line2D([0.045, 0.955], [yy - 0.008] * 2,
                                        transform=fig.transFigure, color="#B9B4A8", lw=0.8))
        zebra = 0
        for ln in pls:
            if ln[0] == "g":
                y -= 0.010
                fig.text(XN, y, ln[1], fontsize=9.8, fontweight="bold", color=NAVY2)
                y -= 0.022
                colhead(y); y -= 0.024
                zebra = 0
            else:
                _, k, h, e = ln
                m = charts[k]
                if zebra % 2 == 0:
                    fig.patches.append(_Rect((0.045, y - 0.0068), 0.91, 0.0222,
                                             color="#F4F1EA", transform=fig.transFigure, zorder=0))
                zebra += 1
                fig.text(XN, y, k.split(" — ")[0], fontsize=8.6, zorder=2)
                fig.text(XA, y, m.get("adm", ""), fontsize=8.2, color="#444", zorder=2)
                fig.text(XO, y, m.get("orbit", ""), fontsize=8.2, color="#444", zorder=2)
                if h + e:
                    r_ = f"ทับ {h} จุด" + (f"  ·  แตะขอบ {e}" if e else "")
                    fig.text(XR, y, r_, fontsize=8.4, ha="right", color=RED2,
                             fontweight="bold", zorder=2)
                    fig.text(XP, y, str(page_no[k]), fontsize=8.6, ha="right",
                             color=NAVY2, fontweight="bold", zorder=2)
                else:
                    fig.text(XR, y, "ไม่พบการทับ/แตะขอบ", fontsize=8.4, ha="right",
                             color=GRN2, zorder=2)
                    fig.text(XP, y, "—", fontsize=8.6, ha="right", color="#999", zorder=2)
                y -= 0.0222
        fig.text(0.5, 0.018, f"— {ip + 1} —", ha="center", fontsize=7, color="#888")
        pdf.savefig(fig); plt.close(fig)
    for i, (k, h, e) in enumerate(ordered_rows, start=1):
        if progress:
            progress(i / len(ordered_rows))
        if not (h + e):
            continue
        fig, _why = collision_fig(k, charts[k], thai, ific)
        if fig is None:
            continue
        fig.text(0.985, 0.006, f"หน้า {page_no[k]}", ha="right", fontsize=6.5, color="#888")
        pdf.savefig(fig); plt.close(fig)
    pdf.close()
    return buf.getvalue()

# ================= ร่างบันทึกข้อความ (.docx) =================

_TH_DIG = str.maketrans("0123456789", "๐๑๒๓๔๕๖๗๘๙")
_TH_MONTH = ["", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
             "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]

def _tn(x):
    return str(x).translate(_TH_DIG)

def build_memo_docx(ific, m97, m912, m912a, m921, m914, memo93=None):
    """ร่างบันทึกข้อความตามฟอร์ม ดบ. (A4, TH SarabunPSK 16, ขอบ 1.5/2/3/2 ซม.,
    เลขหน้าไทย) — ช่องที่แอปไม่มีทางรู้ใส่ (ระบุ) ไว้ให้เติม"""
    import datetime as _dt
    doc = _Docx()
    stn = doc.styles["Normal"]
    stn.font.name = "TH SarabunPSK"; stn.font.size = _Pt(16)
    rpr = stn.element.get_or_add_rPr(); rf = rpr.get_or_add_rFonts()
    for k in ("w:ascii", "w:hAnsi", "w:cs"):
        rf.set(_qn(k), "TH SarabunPSK")
    szcs = _OE("w:szCs"); szcs.set(_qn("w:val"), "32"); rpr.append(szcs)
    sec = doc.sections[0]
    sec.page_width, sec.page_height = _Cm(21.0), _Cm(29.7)
    sec.top_margin, sec.bottom_margin = _Cm(1.5), _Cm(2.0)
    sec.left_margin, sec.right_margin = _Cm(3.0), _Cm(2.0)
    sec.header_distance, sec.footer_distance = _Cm(1.25), _Cm(1.5)
    pgt = _OE("w:pgNumType"); pgt.set(_qn("w:fmt"), "thaiNumbers")
    sec._sectPr.append(pgt)
    hp = sec.header.paragraphs[0]; hp.alignment = _AL.CENTER
    hp.add_run("- ")
    fld = _OE("w:fldSimple"); fld.set(_qn("w:instr"), r"PAGE \* MERGEFORMAT")
    hp._p.append(fld)
    hp.add_run(" -")

    def para(text="", indent=None, align=None, bold=False, size=None, after=4):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = _Pt(after)
        if indent is not None:
            p.paragraph_format.first_line_indent = _Cm(indent)
        if align is not None:
            p.alignment = align
        r = p.add_run(text)
        r.bold = bold
        if size:
            r.font.size = _Pt(size)
        return p

    def _fix_layout(t, widths):
        t.autofit = False
        tbl = t._tbl
        lay = _OE("w:tblLayout"); lay.set(_qn("w:type"), "fixed")
        tbl.tblPr.append(lay)
        for gc, w in zip(tbl.tblGrid.findall(_qn("w:gridCol")), widths):
            gc.set(_qn("w:w"), str(w))
        for row in t.rows:
            seen = set()
            uniq = []
            for c, w in zip(row.cells, widths):
                if id(c._tc) in seen:
                    continue
                seen.add(id(c._tc)); uniq.append((c, w))
            if len(uniq) == 1:
                uniq[0][0].width = _Emu(sum(widths) * 635)
            else:
                for c, w in uniq:
                    c.width = _Emu(w * 635)

    def memo_table(headers, rows, widths):
        t = doc.add_table(rows=1, cols=len(headers))
        t.style = "Table Grid"; t.alignment = _TAL.CENTER
        for c, h, w in zip(t.rows[0].cells, headers, widths):
            c.width = _Emu(w * 635)
            cp = c.paragraphs[0]; cp.alignment = _AL.CENTER
            r = cp.add_run(h); r.bold = True; r.font.size = _Pt(14)
        for row in rows:
            cells = t.add_row().cells
            for j, (c, v) in enumerate(zip(cells, row)):
                c.width = _Emu(widths[j] * 635)
                lines = str(v).split("\n")
                cp = c.paragraphs[0]
                cp.alignment = _AL.CENTER if j < 5 else _AL.LEFT
                cp.add_run(lines[0]).font.size = _Pt(14)
                for ln in lines[1:]:
                    p2 = c.add_paragraph(); p2.alignment = cp.alignment
                    p2.add_run(ln).font.size = _Pt(14)
        _fix_layout(t, widths)
        return t

    today = _dt.date.today()
    ific_th = _tn(ific)
    # ---- หัวบันทึก ----
    p = doc.add_paragraph(); p.paragraph_format.space_after = _Pt(2)
    for gp in ("garuda.png", "fonts/garuda.png"):
        if os.path.exists(gp):
            p.add_run().add_picture(gp, height=_Cm(1.65))
            break
    r = p.add_run("   บันทึกข้อความ"); r.font.size = _Pt(29); r.bold = True
    ht = doc.add_table(rows=3, cols=2)
    ht.alignment = _TAL.CENTER
    hw = (6500, 3276)
    ht.rows[0].cells[0].merge(ht.rows[0].cells[1])
    ht.rows[2].cells[0].merge(ht.rows[2].cells[1])
    ht.rows[0].cells[0].paragraphs[0].add_run(
        "หน่วยงาน   สำนักกิจการดาวเทียมสื่อสาร  (ส่วนประสานงานกิจการดาวเทียม) โทร. ๒๗๑๑")
    ht.rows[1].cells[0].paragraphs[0].add_run("ที่    สทช ๒๐๐๘/(ระบุ)")
    ht.rows[1].cells[1].paragraphs[0].add_run(
        f"วันที่          {_TH_MONTH[today.month]} {_tn(today.year + 543)}")
    ht.rows[2].cells[0].paragraphs[0].add_run(
        "เรื่อง    ความคิดเห็นของประเทศไทยต่อ Administration ที่รับผิดชอบดาวเทียมที่มีแนวโน้ม"
        "ก่อให้เกิดการรบกวนต่อข่ายงานดาวเทียมของประเทศไทยและข่ายสื่อสารภาคพื้นโลก "
        f"ตาม BR IFIC หมายเลข {ific_th}")
    _fix_layout(ht, list(hw))
    para("เรียน  ผชพ. (ระบุ)", after=6)

    # ---- ความเป็นมา ----
    para("ความเป็นมา", indent=2.5)
    k = 1
    para(f"{_tn(k)}. (ระบุความเป็นมา เช่น หนังสือของบริษัทผู้ร้องขอ เลขที่/ลงวันที่ "
         "และความประสงค์ให้แจ้งทักท้วง)", indent=2.5)
    para(f"ข้อเท็จจริง ตาม BR IFIC หมายเลข {ific_th}", indent=2.5)
    para(f"ดบ. ตรวจสอบ BR IFIC หมายเลข {ific_th} แล้ว ขอเรียนดังนี้",
         indent=1.3, align=_AL.JUSTIFY)

    W7 = [1980, 1134, 709, 1275, 851, 2551, 1276]
    H7 = ["ข่ายงานดาวเทียม", "Nominal\nLongitude", "BR IFIC", "Special\nsection",
          "ประเทศ", "ข่ายงานดาวเทียม\nที่ได้รับผลกระทบ", "ข้อบังคับวิทยุ\nที่เกี่ยวข้อง"]
    refs = {}

    def _rows_std(df, impact_col):
        out = []
        for d in df.to_dict("records"):
            out.append([d["ข่ายงานดาวเทียม"], d["Nominal Longitude"], _tn(d["BR IFIC"]),
                        d["Special section"], d["ประเทศ"], d.get(impact_col, ""),
                        d.get("มาตรา", d.get("ข้อบังคับวิทยุ", ""))])
        return out

    groups = []
    if len(m97):
        groups.append(("9.7", "ข่ายงานดาวเทียมของประเทศอื่น ที่ต้องประสานงานคลื่นความถี่"
                       "กับข่ายงานดาวเทียมของประเทศไทย ภายใต้ข้อบังคับวิทยุ มาตรา 9.7 มีดังนี้",
                       H7, W7, _rows_std(m97, "ข่ายงานดาวเทียมที่ได้รับผลกระทบ"), None))
    r12 = (_rows_std(m912, "ข่ายงานดาวเทียมที่ได้รับผลกระทบ") if len(m912) else []) + \
          (_rows_std(m912a, "ข่ายงานดาวเทียมที่ได้รับผลกระทบ") if len(m912a) else [])
    if r12:
        r12.sort(key=lambda x: int(sec_num(x[3]) or 999999))
        groups.append(("9.12AB", "ข่ายงานดาวเทียมของประเทศอื่น ที่ต้องประสานงานคลื่นความถี่"
                       "กับข่ายงานดาวเทียมของประเทศไทย ภายใต้ข้อบังคับวิทยุ มาตรา 9.12 "
                       "และ 9.12A มีดังนี้", H7, W7, r12, None))
    m921c = m921[m921["มาตรา"] == "9.21/C"] if len(m921) else m921
    for _pv, _scope in (("9.21/A", "GSO"), ("9.21/B", "NGSO")):
        sub = m921[m921["มาตรา"] == _pv] if len(m921) else m921
        if not len(sub):
            continue
        rows = []
        for d in sub.to_dict("records"):
            imp = d["ข่ายไทยที่ทับ (ทับจริง)"]
            if str(d.get("แตะขอบ (0 MHz)", "")).strip() not in ("", "—"):
                imp = f"{imp}\nแตะขอบ: {d['แตะขอบ (0 MHz)']}"
            rows.append([d["ข่ายงานดาวเทียม"], d["Nominal Longitude"], _tn(d["BR IFIC"]),
                         d["Special section"], d["ประเทศ"], imp, d["มาตรา"]])
        groups.append((_pv, "ข่ายงานดาวเทียม GSO และ NGSO ของประเทศอื่น ที่ต้องประสานงาน"
                       f"คลื่นความถี่กับข่ายงานดาวเทียม {_scope} ของประเทศไทย "
                       f"ภายใต้ข้อบังคับวิทยุ มาตรา {_pv} มีดังนี้", H7, W7, rows, None))
    if len(m921c):
        H6 = H7[:5] + ["ย่านความถี่\n(MHz)", H7[6]]
        rows = []
        for d in m921c.to_dict("records"):
            band = d["ย่านความถี่ (MHz)"]
            svc = str(d.get("บริการภาคพื้นไทย (NFAT)", "")).strip()
            if svc and svc not in ("—", "(ไม่ได้แนบตาราง NFAT)"):
                band = f"{band}\n[NFAT: {svc}]"
            rows.append([d["ข่ายงานดาวเทียม"], d["Nominal Longitude"], _tn(d["BR IFIC"]),
                         d["Special section"], d["ประเทศ"], band, d["มาตรา"]])
        groups.append(("9.21C", "ข่ายงานดาวเทียมของประเทศอื่น ที่ต้องประสานงานคลื่นความถี่"
                       "กับข่ายสื่อสารภาคพื้นโลกของประเทศไทย ตามเชิงอรรถระหว่างประเทศของ"
                       "ย่านความถี่ที่เกี่ยวข้อง ภายใต้ข้อบังคับวิทยุ มาตรา 9.21/C มีดังนี้",
                       H6, W7, rows,
                       "หมายเหตุ (ระบุเชิงอรรถของย่าน เช่น 5.218/5.254 และจำนวนหน่วยงาน"
                       "ผู้ใช้งานคลื่นความถี่ของไทยตามเอกสารแนบ)"))
    if len(m914):
        H14 = ["ข่ายงานดาวเทียม", "Nominal\nLongitude", "Special\nsection", "ประเทศ",
               "ข้อบังคับวิทยุ\nที่เกี่ยวข้อง", "ผลการพิจารณาตาม RR\nArticle 9"]
        W14 = [1900, 1300, 1300, 850, 1300, 3126]
        rows = []
        for (nm_, orb, sec_, adm_), g in m914.groupby(
                ["ข่ายงานดาวเทียม", "วงโคจร", "Special section", "ประเทศ"], sort=False):
            res = []
            for d in g.to_dict("records"):
                v = str(d["ผลการพิจารณา"])
                seg = d["ส่วนที่เทียบ (MHz)"]
                if v.startswith("เกิน"):
                    res.append(f"• เกินเกณฑ์: {seg} MHz (margin {d['margin (dB)']} dB)")
                elif "footnote" in v:
                    res.append(f"• ไม่มีเกณฑ์ (ตรวจ footnote): {seg} MHz")
                elif v.startswith("ไม่เกิน"):
                    res.append(f"• ไม่เกินเกณฑ์: {seg} MHz")
                elif v.startswith("มีเฉพาะขาขึ้น"):
                    res.append(f"• ขาขึ้น (R): {d['ย่านที่ตีพิมพ์ (MHz)']}")
                else:
                    res.append(f"• {seg} MHz: {v}")
            rows.append([nm_, orb, sec_, adm_, "9.14", "\n".join(res)])
        groups.append(("9.14", "ข่ายงานดาวเทียมของประเทศอื่น ที่ต้องพิจารณาการแจ้งทักท้วง"
                       "เพื่อคุ้มครองบริการภาคพื้นของประเทศไทย ภายใต้ข้อบังคับวิทยุ มาตรา 9.14 "
                       "โดยผลการคำนวณ pfd เทียบเกณฑ์ Annex 1 ของ Appendix 5 มีดังนี้",
                       H14, W14, rows,
                       "หมายเหตุ การแจ้งทักท้วงตามมาตรา 9.11 และ 9.14 หากผลการคำนวณ PFD fail "
                       "จะแจ้งทุกกรณี หากผลการคำนวณ PFD pass จะแจ้งกรณีที่มีผู้ใช้งานจริง "
                       "หากผลการคำนวณ pass และไม่มีผู้ใช้งานจริง จะไม่แจ้งทักท้วงแม้ว่าจะมี"
                       "การใช้งานคลื่นความถี่ตามตารางการกำหนดคลื่นความถี่แห่งชาติก็ตาม"))
    if memo93 is not None and len(memo93):
        H93 = ["ข่ายงานดาวเทียม", "Nominal\nLongitude", "Special\nsection", "ประเทศ",
               "ข่ายงานดาวเทียม\nที่ได้รับผลกระทบ (S)", "T"]
        W93 = [1900, 1134, 1600, 850, 3816, 476]
        rows = []
        for d in memo93.to_dict("records"):
            s = str(d.get("ข่ายไทยที่กระทบ (S)", "")).strip()
            e = str(d.get("แตะขอบ (0 MHz)", "")).strip()
            if e and e != "—":
                s = (s + "\n" if s and s != "—" else "") + f"แตะขอบ: {e}"
            rows.append([d["ข่ายงานดาวเทียม"], d["Nominal Longitude"],
                         d["Special section"], d["ประเทศ"], s,
                         d.get("T (มีขาลง)", "")])
        groups.append(("93", "ข่ายงานดาวเทียม GSO และ NGSO ของประเทศอื่น ที่ต้อง"
                       "ประสานงานคลื่นความถี่กับประเทศไทย ภายใต้ข้อบังคับวิทยุ มาตรา 9.3 "
                       "(S: Interference to the space services  T: Interference to the "
                       "terrestrial services และ E: Excluding territory) มีดังนี้",
                       H93, W93, rows,
                       "หมายเหตุ คอลัมน์ (S) = ข่ายดาวเทียมไทยที่ย่านความถี่ทับ จากการเทียบ"
                       "ของ ปบ. · คอลัมน์ T ใส่ X เมื่อข่ายนั้นมี emission ขาลง "
                       "(มีโอกาสกระทบบริการภาคพื้น) — MDB ไม่ได้ให้ flag S/T ของ BR มา "
                       "โปรดตรวจทานกับตัว Special Section"))

    for key, lead, hd, wd, rows, note in groups:
        k += 1
        refs[key] = k
        para(f"{_tn(k)}. {lead}", indent=2.5, align=_AL.JUSTIFY)
        memo_table(hd, rows, wd)
        if note:
            para(note, size=14, after=6)
        else:
            para("", after=2)

    # ---- ข้อเสนอ ----
    para("ข้อเสนอเพื่อพิจารณา", indent=1.8)
    k += 1
    para(f"{_tn(k)}. ดบ. พิจารณาแล้ว ขอเสนอความเห็นดังนี้", indent=2.5)
    prop = {
        "9.7": "เห็นควรแจ้งทักท้วงข่ายงานดาวเทียมของประเทศอื่น ตามข้อ {r}. เนื่องจาก"
               "ข่ายงานดาวเทียมของประเทศไทยอาจได้รับการรบกวน (ตรวจทาน/เพิ่มเหตุผลผู้ร้องขอ)",
        "9.12AB": "เห็นควรแจ้งทักท้วงข่ายงานดาวเทียมของประเทศอื่น ตามข้อ {r}. เนื่องจาก"
                  "ข่ายงานดาวเทียมของประเทศไทยอาจได้รับการรบกวน",
        "93": "เห็นควรมีหนังสือแจ้งความคิดเห็นของประเทศไทยไปยัง administration "
              "ที่เกี่ยวข้อง ตามข้อ {r}. ภายใต้ข้อบังคับวิทยุ มาตรา 9.3 เนื่องจาก"
              "ข่ายงานดาวเทียมของประเทศไทยอาจได้รับการรบกวน",
        "9.21/A": "เห็นควรแจ้งทักท้วงข่ายงานดาวเทียมของประเทศอื่น ตามข้อ {r}. เนื่องจาก"
                  "ข่ายงานดาวเทียม GSO ของประเทศไทยอาจได้รับการรบกวน",
        "9.21/B": "เห็นควรแจ้งทักท้วงข่ายงานดาวเทียมของประเทศอื่น ตามข้อ {r}. เนื่องจาก"
                  "ข่ายงานดาวเทียม NGSO ของประเทศไทยอาจได้รับการรบกวน",
        "9.21C": "เห็นควรแจ้งทักท้วงข่ายงานดาวเทียมของประเทศอื่น ตามข้อ {r}. เนื่องจาก"
                 "ข่ายสื่อสารภาคพื้นโลกของประเทศไทยอาจได้รับการรบกวน",
        "9.14": "เห็นควรแจ้งทักท้วงข่ายงานดาวเทียมของประเทศอื่น ตามข้อ {r}. เฉพาะรายการ"
                "ที่ผลการคำนวณ pfd เกินเกณฑ์ หรือไม่เกินเกณฑ์แต่มีผู้ใช้งานจริง "
                "ตามแนวทางในหมายเหตุท้ายข้อ {r}.",
    }
    for i, (key, *_rest) in enumerate(groups, start=1):
        para(f"{_tn(i)}) {prop[key].format(r=_tn(refs[key]))}", indent=3.3, align=_AL.JUSTIFY)
    para("จึงเรียนมาเพื่อโปรดพิจารณา หากเห็นชอบตามที่ ดบ. มีความเห็น ขอได้โปรดลงนาม"
         "ในหนังสือแจ้งหน่วยงานที่รับผิดชอบข่ายงานดาวเทียมของประเทศอื่น และอนุญาตให้ ดบ. "
         "จัดส่งข้อมูลแจ้งทักท้วงทางอิเล็กทรอนิกส์และสำเนาให้สำนักงานวิทยุคมนาคม "
         "ภายในวันที่ (ระบุวันครบกำหนดส่งความเห็น) ต่อไปด้วย จะขอบคุณยิ่ง",
         indent=2.0, align=_AL.JUSTIFY, after=24)
    sig = doc.add_paragraph(); sig.alignment = _AL.CENTER
    sig.paragraph_format.left_indent = _Cm(7.5)
    sig.add_run("((ระบุชื่อผู้เสนอ))\n(ระบุตำแหน่ง)")
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

def build_crc(mdb_path, ific, thai, allowed_secnums=None, allowed_ids=None, nfat=None):
    """allowed_secnums/ids มาจาก portal CSV (ขอบเขตทางการ) — รายการนอกขอบเขตจะถูกแยกออกมา
    ใน excluded (ไม่ทิ้งเงียบ); nfat = ตารางกำหนดคลื่นความถี่แห่งชาติ (บริการภาคพื้นไทย)"""
    tr = mdb_table(mdb_path, "tr_aff_ntw")
    pv = mdb_table(mdb_path, "provn")
    com = mdb_table(mdb_path, "com_el")
    pub = mdb_table(mdb_path, "pub_ssn")
    grp = mdb_table(mdb_path, "grp")
    em = mdb_table(mdb_path, "emiss")
    sb = mdb_table(mdb_path, "s_beam")
    orb = mdb_table(mdb_path, "orbit")
    cinfo = {c["ntc_id"]: c for c in com}
    ssn = ssn_map(pub, "CR/C")
    gmap = {g["grp_id"]: g for g in grp}
    # ข้อมูลกำลังส่งสำหรับ pfd 9.14 (ฟิลด์พิสูจน์แล้ว: emiss.pwr_ds_max=C8a2/C8b2,
    # s_beam.gain=B3a1, s_beam.pwr_max_4k/1m=EIRP ต่อ BW อ้างอิง, orbit.perig_km=A4b4e)
    embyg = defaultdict(list)
    for e in em:
        if e.get("pwr_ds_max"):
            embyg[e["grp_id"]].append(float(e["pwr_ds_max"]))
    sbk = {(b["ntc_id"], b["beam_name"], b["emi_rcp"]): b for b in sb}
    perig = {}
    for o in orb:
        if o.get("perig_km"):
            n = o["ntc_id"]
            perig[n] = min(perig.get(n, 1e9), float(o["perig_km"]))

    def in_scope(sec, ids):
        if allowed_secnums is None:
            return True
        return (sec_num(sec) in allowed_secnums) or bool(ids & (allowed_ids or set()))

    # ---- Engine B: 9.7 (รวม -CI/-CII) + 9.12A จาก tr_aff_ntw ----
    agg = defaultdict(lambda: {"thai": set(), "ids": set(), "flags": []})
    for r in tr:
        if r.get("adm") == "THA" and r.get("wic_no") == ific:
            prov = crc_prov_norm(r["coord_prov"])
            sec = ssn.get(r["ntc_id"], f"(ntc {r['ntc_id']})")
            a = agg[(prov, sec)]
            a["ids"].add(r["ntc_id"])
            a["thai"].add(thai_label(r["sat_name"], r.get("long_nom", "")))
            a["flags"].append((r["sat_name"], r.get("f_cause", ""), r.get("f_rec", "")))

    def foreign_of(ids):
        for i in sorted(ids):
            c = cinfo.get(i)
            if c and c.get("sat_name"):
                return c["sat_name"], fmt_lon(c.get("long_nom", "")), c.get("adm", "")
        return f"ID {sorted(ids)[0]}", "?", "?"

    mem_b, flags_rows, excluded = {"9.7": [], "9.12": [], "9.12A": [], "AP30": []}, [], []
    for (prov, sec), a in sorted(agg.items(), key=lambda x: (x[0][0], x[0][1])):
        name, lon, adm = foreign_of(a["ids"])
        if not in_scope(sec, a["ids"]):
            excluded.append({"มาตรา": prov, "Special section": sec, "ข่ายงานดาวเทียม": name,
                             "เหตุผล": "ไม่อยู่ในรายการตีพิมพ์ทางการ (portal) ของรอบนี้"})
            continue
        tgt = mem_b.setdefault(prov, [])
        tgt.append({
            "ข่ายงานดาวเทียม": name, "Nominal Longitude": lon, "BR IFIC": ific,
            "Special section": sec, "ประเทศ": adm,
            "ข่ายงานดาวเทียมที่ได้รับผลกระทบ": nl(thai_sort(a["thai"])),
            "ข้อบังคับวิทยุ": prov,
        })
        for tname, fc, fr in a["flags"]:
            flags_rows.append({"ข่ายต่างชาติ": name, "มาตรา": prov, "Special section": sec,
                               "ข่ายไทย": tname, "f_cause": fc, "f_rec": fr})

    # ---- Engine A: 9.21/* + 9.14 จาก provn + grp (กรอง wic) ----
    bands = defaultdict(lambda: defaultdict(set))  # (prov, sec) -> (fmin,fmax) -> {emi}
    ids_a = defaultdict(set)
    # ข้อมูลกำลังส่งแย่สุดต่อย่าน (เฉพาะ emission ขาลง E) สำหรับ pfd 9.14:
    # (prov,sec) -> (fmin,fmax) -> {"e4":(EIRP,pd,gain,beam,วิธี), "e1":(...)}
    pinfo = defaultdict(dict)
    for p in pv:
        if p.get("adm") == "THA" and (p["coord_prov"].startswith("9.21") or p["coord_prov"] == "9.14"):
            g = gmap.get(p["grp_id"])
            if g and g.get("wic_no") == ific and g.get("freq_min"):
                sec = ssn.get(g["ntc_id"], f"(ntc {g['ntc_id']})")
                key = (p["coord_prov"], sec)
                bk = (float(g["freq_min"]), float(g["freq_max"]))
                bands[key][bk].add(g.get("emi_rcp", ""))
                ids_a[key].add(g["ntc_id"])
                if p["coord_prov"] == "9.14" and g.get("emi_rcp") == "E":
                    b = sbk.get((g["ntc_id"], g.get("beam_name", ""), "E"), {})
                    gain = float(b["gain"]) if b.get("gain") else None
                    pdm = max(embyg.get(g["grp_id"], [-1e9]))
                    slot = pinfo[key].setdefault(bk, {"e4": None, "e1": None})
                    for fld, addb in (("e4", 36.02), ("e1", 60.0)):
                        mdbf = b.get("pwr_max_4k" if fld == "e4" else "pwr_max_1m")
                        if mdbf:
                            cand = (float(mdbf), pdm if pdm > -1e8 else None, gain,
                                    g.get("beam_name", ""), "ฟิลด์ pwr_max")
                        elif pdm > -1e8 and gain is not None:
                            cand = (pdm + addb + gain, pdm, gain, g.get("beam_name", ""), "คำนวณ")
                        else:
                            cand = None
                        if cand and (slot[fld] is None or cand[0] > slot[fld][0]):
                            slot[fld] = cand

    # ฐานไทยแยกตามวงโคจร สำหรับ 9.21/A (GSO) และ 9.21/B (NGSO) — อิงชื่อในฐาน:
    # ลงท้าย "(NGSO)" = NGSO, ลงท้ายตำแหน่ง "(xxE)" = GSO
    thai_gso = thai_ngso = None
    if thai is not None:
        _is_ngso = thai["network_name"].astype(str).str.upper().str.contains(r"\(\s*NGSO\s*\)")
        thai_gso, thai_ngso = thai[~_is_ngso], thai[_is_ngso]

    mem_a = {"9.21": [], "9.14": []}
    det_a, pfd_lines, att6 = [], [], []
    for (prov, sec), bset in sorted(bands.items(), key=lambda x: (x[0][0], x[0][1])):
        name, lon, adm = foreign_of(ids_a[(prov, sec)])
        if not in_scope(sec, ids_a[(prov, sec)]):
            excluded.append({"มาตรา": prov, "Special section": sec, "ข่ายงานดาวเทียม": name,
                             "เหตุผล": "ไม่อยู่ในรายการตีพิมพ์ทางการ (portal) ของรอบนี้"})
            continue

        # ---------- 9.21: แยกตามนิยามของ Special Section (เชิงอรรถ CR/C) ----------
        # 9.21/A = ข่าย GSO ไทย | 9.21/B = ข่าย NGSO ไทย | 9.21/C = สถานีภาคพื้นไทย
        if prov != "9.14":
            band_txt = "\n".join(
                f"{a:g} – {b:g} MHz [{'/'.join(sorted(e for e in emis if e))}]"
                for (a, b), emis in sorted(bset.items()))
            if prov == "9.21/C":
                svc_txt = "\n".join(f"{a:g}–{b:g}: {nfat_services(nfat, a, b)}"
                                    for (a, b) in sorted(bset)) if nfat is not None \
                          else "(ไม่ได้แนบตาราง NFAT)"
                mem_a["9.21"].append({
                    "ข่ายงานดาวเทียม": name, "Nominal Longitude": lon, "BR IFIC": ific,
                    "Special section": sec, "ประเทศ": adm,
                    "ย่านความถี่ (MHz)": band_txt,
                    "ข่ายไทยที่ทับ (ทับจริง)": "— กรณีสถานีภาคพื้น (ไม่เทียบข่ายดาวเทียม)",
                    "แตะขอบ (0 MHz)": "—",
                    "บริการภาคพื้นไทย (NFAT)": svc_txt,
                    "หมายเหตุ": "เกณฑ์ตาม footnote ของย่าน (AP5 Table 5-1) — ตรวจ footnote ประกอบ",
                    "มาตรา": prov,
                })
            else:
                tdf = (thai_gso if prov == "9.21/A"
                       else thai_ngso if prov == "9.21/B" else thai)
                real, edge = set(), set()
                if tdf is not None:
                    for (a, b) in bset:
                        for t in tdf.itertuples(index=False):
                            lo = max(a, t.freq_start_mhz); hi = min(b, t.freq_end_mhz)
                            if hi - lo > 0:
                                real.add(t.network_name)
                            elif hi - lo == 0:
                                edge.add(t.network_name)
                    edge -= real
                note = ("เทียบเฉพาะข่าย GSO ไทย (ตามนิยาม 9.21/A)" if prov == "9.21/A"
                        else "เทียบเฉพาะข่าย NGSO ไทย (ตามนิยาม 9.21/B)" if prov == "9.21/B" else "")
                mem_a["9.21"].append({
                    "ข่ายงานดาวเทียม": name, "Nominal Longitude": lon, "BR IFIC": ific,
                    "Special section": sec, "ประเทศ": adm,
                    "ย่านความถี่ (MHz)": band_txt,
                    "ข่ายไทยที่ทับ (ทับจริง)": nl(thai_sort(real)) if tdf is not None else "(ต้องมีฐานไทย)",
                    "แตะขอบ (0 MHz)": nl(thai_sort(edge)) if tdf is not None else "",
                    "บริการภาคพื้นไทย (NFAT)": "—",
                    "หมายเหตุ": note,
                    "มาตรา": prov,
                })
            for (a, b), emis in sorted(bset.items()):
                det_a.append({"ข่ายต่างชาติ": name, "มาตรา": prov, "Special section": sec,
                              "f_min": a, "f_max": b,
                              "ทิศทาง": "/".join(sorted(e for e in emis if e))})
            continue

        # ---------- 9.14: คัดกรอง pfd เทียบเกณฑ์ AP5 (คุ้มครองบริการภาคพื้น) ----------
        c0 = next((cinfo[i] for i in sorted(ids_a[(prov, sec)]) if i in cinfo), {})
        lon_raw = (c0.get("long_nom") or "").strip()
        if lon_raw:
            kind, par = "GSO", float(lon_raw)
            orbit_txt = f"GSO {fmt_lon(lon_raw)}"
        else:
            pk = min((perig[i] for i in ids_a[(prov, sec)] if i in perig), default=None)
            kind, par = "NGSO", pk
            orbit_txt = f"NGSO (perigee {pk:g} km)" if pk else "NGSO (ไม่พบ perigee)"
        gso = kind == "GSO"
        vis = _gso_points(par) if gso else None
        pfd_lines += ["", f"━━━ {name} · {orbit_txt} · {sec} · {adm} ━━━"]
        if gso:
            if vis:
                evals = [p[0] for p in vis]
                pfd_lines.append(f"หน้าต่างมุมเงยจากอาณาเขตไทย: {min(evals):.1f}° – {max(evals):.1f}° "
                                 f"(กวาดกรอบพิกัดจุดสุดเขตแดน 4 ทิศ ทีละ 0.25°)")
            else:
                pfd_lines.append("ดาวเทียมอยู่ใต้ขอบฟ้าของไทยทั้งประเทศ (ไม่มีจุดใดมองเห็น)")

        for (f1, f2), emis in sorted(bset.items()):
            dirs = "/".join(sorted(e for e in emis if e))
            base = {"ข่ายงานดาวเทียม": name, "วงโคจร": orbit_txt, "BR IFIC": ific,
                    "Special section": sec, "ประเทศ": adm,
                    "ย่านที่ตีพิมพ์ (MHz)": f"{f1:g} – {f2:g} [{dirs}]"}
            segs, un = _th_segments(f1, f2, gso)
            pfd_lines.append(f"ย่านตีพิมพ์ {f1:g} – {f2:g} MHz [{dirs}]")

            if "E" not in emis:
                mem_a["9.14"].append({**base, "ส่วนที่เทียบ (MHz)": "—", "เกณฑ์ ณ θ* [ที่มา]": "—",
                                      "pfd คำนวณ": "—", "margin (dB)": "—",
                                      "ผลการพิจารณา": "มีเฉพาะขาขึ้น (R) — ไม่เข้ากรณี pfd ขาลง",
                                      "บริการภาคพื้นไทย (NFAT)": nfat_services(nfat, f1, f2) if nfat is not None else "—",
                                      "มาตรา": "9.14"})
                att6.append({"ข่ายงานดาวเทียม": name, "Frequency (MHz)": f"{f1:g} – {f2:g}",
                             "Power density dB(W/Hz)": "—", "Calculated pfd": "—",
                             "Max. pfd from Annex 1 of Appendix 5 dB(W/m²)": "—",
                             "Pass/Fail": "— (ขาขึ้น R)"})
                pfd_lines.append("  ▸ มีเฉพาะ emission ขาขึ้น (R) — ไม่เข้ากรณี pfd ขาลง")
                continue

            info = pinfo.get((prov, sec), {}).get((f1, f2), {"e4": None, "e1": None})
            for a, b, bw, P, r, srcname in segs:
                bwtxt = "4 kHz" if bw == "4kHz" else "1 MHz"
                addb = 36.02 if bw == "4kHz" else 60.0
                tup = info["e4"] if bw == "4kHz" else info["e1"]
                seg_lbl = f"{a:g} – {b:g}"
                if gso and not vis:
                    mem_a["9.14"].append({**base, "ส่วนที่เทียบ (MHz)": seg_lbl,
                                          "เกณฑ์ ณ θ* [ที่มา]": f"[{srcname}]",
                                          "pfd คำนวณ": "—", "margin (dB)": "—",
                                          "ผลการพิจารณา": "ไม่เกิน — ใต้ขอบฟ้าไทยทั้งประเทศ",
                                          "บริการภาคพื้นไทย (NFAT)": nfat_services(nfat, a, b) if nfat is not None else "—",
                                          "มาตรา": "9.14"})
                    att6.append({"ข่ายงานดาวเทียม": name, "Frequency (MHz)": seg_lbl,
                                 "Power density dB(W/Hz)": "—", "Calculated pfd": "—",
                                 "Max. pfd from Annex 1 of Appendix 5 dB(W/m²)": f"[{srcname}]",
                                 "Pass/Fail": "Pass (ใต้ขอบฟ้าไทย)"})
                    continue
                if tup is None or (kind == "NGSO" and par is None):
                    mem_a["9.14"].append({**base, "ส่วนที่เทียบ (MHz)": seg_lbl,
                                          "เกณฑ์ ณ θ* [ที่มา]": f"P={P}, r={r if r is not None else '—'} [{srcname}]",
                                          "pfd คำนวณ": "—", "margin (dB)": "—",
                                          "ผลการพิจารณา": "ข้อมูลใน MDB ไม่ครบ → ตรวจจากเอกสารข่าย",
                                          "บริการภาคพื้นไทย (NFAT)": nfat_services(nfat, a, b) if nfat is not None else "—",
                                          "มาตรา": "9.14"})
                    att6.append({"ข่ายงานดาวเทียม": name, "Frequency (MHz)": seg_lbl,
                                 "Power density dB(W/Hz)": "—", "Calculated pfd": "ตรวจจากเอกสารข่าย",
                                 "Max. pfd from Annex 1 of Appendix 5 dB(W/m²)": f"{P} [{srcname}]",
                                 "Pass/Fail": "—"})
                    pfd_lines.append(f"  ▸ ส่วน {seg_lbl} MHz: ข้อมูลกำลังส่ง/วงโคจรไม่ครบใน MDB — ไม่คำนวณ (ห้ามเดา)")
                    continue
                E, pwr, gn, beam, how = tup
                best = pfd_worst(kind, par, E, P, r)
                m, ts, d, pt = best
                pfdv = E - _spread_db(d)
                Tv = _T(ts, P, r)
                verdict = ("เกิน → เข้าเกณฑ์แจ้งทักท้วง" if m > 0 else "ไม่เกิน → ไม่ต้องแจ้ง")
                if abs(m) <= 1.0:
                    verdict += " (ปริ่มเส้น — พิจารณา SSM/Rec. M.1143)"
                mem_a["9.14"].append({**base, "ส่วนที่เทียบ (MHz)": seg_lbl,
                                      "เกณฑ์ ณ θ* [ที่มา]": f"{Tv:.1f} /{bwtxt} ที่ θ={ts:.1f}° [{srcname}]",
                                      "pfd คำนวณ": f"{pfdv:.1f} /{bwtxt} ที่ θ={ts:.1f}°",
                                      "margin (dB)": f"{m:+.1f}",
                                      "ผลการพิจารณา": verdict,
                                      "บริการภาคพื้นไทย (NFAT)": nfat_services(nfat, a, b) if nfat is not None else "—",
                                      "มาตรา": "9.14"})
                att6.append({"ข่ายงานดาวเทียม": name, "Frequency (MHz)": seg_lbl,
                             "Power density dB(W/Hz)": (f"{pwr:.1f}" if pwr is not None else "—"),
                             "Calculated pfd": f"{pfdv:.1f} /{bwtxt}\n(ที่มุม {ts:.1f}°)",
                             "Max. pfd from Annex 1 of Appendix 5 dB(W/m²)":
                                 (f"{Tv:.0f} /{bwtxt}\n(ทุกมุม, §1.1)" if r is None
                                  else f"{Tv:.1f} /{bwtxt}\n(ที่มุม {ts:.1f}°, Table 5-2)"),
                             "Pass/Fail": ("Fail" if m > 0 else "Pass") + (" *ปริ่มเส้น" if abs(m) <= 1.0 else "")})
                dat = (f"pwr_ds_max = {pwr:.1f} dBW/Hz (C8a2/C8b2)" if pwr is not None else "pwr_ds_max: —")
                gtx = (f"gain = {gn:.1f} dBi (B3a1, บีม {beam})" if gn is not None else f"บีม {beam}")
                pfd_lines.append(f"  ▸ ส่วน {seg_lbl} MHz — เกณฑ์ {srcname} ({kind}): P = {P}, "
                                 f"r = {r if r is not None else 'คงที่ทุกมุม'}, อ้างอิงต่อ {bwtxt}")
                if how == "คำนวณ" and pwr is not None and gn is not None:
                    pfd_lines.append(f"      ขั้น 1  EIRP = pwr_ds_max + 10log₁₀(B_ref) + G = "
                                     f"{pwr:.1f} + {addb:.2f} + {gn:.1f} = {E:.1f} dBW/{bwtxt}")
                else:
                    pfd_lines.append(f"      ขั้น 1  EIRP/{bwtxt} = {E:.1f} dBW — อ่านตรงจากฟิลด์ s_beam.pwr_max "
                                     f"(= pwr_ds_max + 10log₁₀(B_ref) + G; {dat} | {gtx})")
                loc = f" (จุด {pt[0]:.2f}N, {pt[1]:.2f}E บนกรอบไทย)" if pt else ""
                dform = ("d(θ) = Rₑ[√(((Rₑ+h)/Rₑ)² − cos²θ) − sinθ], h = perigee" if kind == "NGSO"
                         else "d จากกฎโคไซน์: cosψ = cosφ·cosΔλ, d = √(Rₑ²+R²−2RₑR·cosψ)")
                pfd_lines.append(f"      ขั้น 2  ระยะทาง: {dform} → จุดแย่สุด θ* = {ts:.1f}°, d = {d:,.0f} km{loc}")
                pfd_lines.append(f"      ขั้น 3  pfd(θ*) = EIRP − 10log₁₀(4πd²) = {E:.1f} − {_spread_db(d):.2f} "
                                 f"= {pfdv:.1f} dBW/m²·{bwtxt}")
                pfd_lines.append(f"      ขั้น 4  เกณฑ์ T(θ*) = {('P = ' + format(Tv, '.1f') + ' (คงที่ทุกมุม)') if r is None else ('P + r(θ−5) = ' + str(P) + ' + ' + str(r) + f'×({ts:.1f}−5) = ' + format(Tv, '.1f') if 5 < ts <= 25 else ('P = ' + format(Tv, '.1f') + ' (θ≤5°)' if ts <= 5 else 'P + 20r = ' + str(P) + ' + ' + format(20 * r, 'g') + ' = ' + format(Tv, '.1f')))} dBW/m²·{bwtxt}")
                pfd_lines.append(f"      ขั้น 5  margin = pfd − T = {pfdv:.1f} − ({Tv:.1f}) = {m:+.1f} dB → {verdict}")
                if nfat is not None:
                    pfd_lines.append(f"      บริการภาคพื้นไทยในย่านนี้ (NFAT): {nfat_services(nfat, a, b)}")
                if kind == "NGSO":
                    chk = []
                    for tc in (5.0, 25.0, 90.0):
                        dc = _slant_km(tc, par)
                        chk.append(f"θ={tc:g}°: {E - _spread_db(dc) - _T(tc, P, r):+.1f}")
                    pfd_lines.append(f"      (ตรวจทาน margin รายมุม — {' | '.join(chk)} dB)")
            for a, b in un:
                mem_a["9.14"].append({**base, "ส่วนที่เทียบ (MHz)": f"{a:g} – {b:g}",
                                      "เกณฑ์ ณ θ* [ที่มา]": "ไม่มีใน Table 5-2 / §1.1",
                                      "pfd คำนวณ": "—", "margin (dB)": "—",
                                      "ผลการพิจารณา": "ไม่มีเกณฑ์ → ตรวจ footnote ของย่าน",
                                      "บริการภาคพื้นไทย (NFAT)": nfat_services(nfat, a, b) if nfat is not None else "—",
                                      "มาตรา": "9.14"})
                att6.append({"ข่ายงานดาวเทียม": name, "Frequency (MHz)": f"{a:g} – {b:g}",
                             "Power density dB(W/Hz)": "—", "Calculated pfd": "—",
                             "Max. pfd from Annex 1 of Appendix 5 dB(W/m²)": "ไม่มีเกณฑ์กำหนด",
                             "Pass/Fail": "—*ตรวจ footnote"})
                pfd_lines.append(f"  ▸ ส่วน {a:g} – {b:g} MHz: ไม่มีเกณฑ์ใน Table 5-2/§1.1 → ตรวจ footnote (ห้ามยืมย่านข้างเคียง)")
                if nfat is not None:
                    pfd_lines.append(f"      บริการภาคพื้นไทยในย่านนี้ (NFAT): {nfat_services(nfat, a, b)}")

    if pfd_lines:
        pfd_lines += ["", "─── สมมติฐานการคำนวณ (เหมือนคู่มือฉบับตำรา) ───",
                      "• worst-case ต่อสถานีอวกาศหนึ่งดวง: กำลังสูงสุด + gain สูงสุด + ระยะใกล้สุด (NGSO ใช้ perigee) + free-space",
                      "• เกณฑ์: RR Appendix 5 Annex 1 — Table 5-2 (Rev.WRC-19, แถว All other cases; MSS) และ §1.1 (ต่ำกว่า 1 GHz); "
                      "สูตรมุมตาม NOTE 2: T=P (0–5°), P+r(θ−5) (5–25°), P+20r (25–90°)",
                      "• GSO: กวาดกรอบพิกัดไทยจากจุดสุดเขตแดน 4 ทิศ (5.610–20.468N, 97.350–105.617E) — เข้มกว่าจริง ~0.2 dB (ทิศปลอดภัย)",
                      "• NGSO: กวาดมุมตกกระทบ 0–90° ทีละ 0.1°",
                      "• ผล 'เกิน' = เข้าเกณฑ์ต้องประสานงาน/มีสิทธิ์ทักท้วง มิใช่คำยืนยันว่ารบกวนจริง; ผลปริ่มเส้น (±1 dB) ควรวิเคราะห์ต่อด้วย SSM ตาม Rec. ITU-R M.1143 (AP5 §1.2.3.2)"]

    if pfd_lines:
        pfd_lines = [
            "สูตรที่ใช้ (ที่มา: คู่มือการคำนวณ pfd ฉบับตำรา — RR Appendix 5 Annex 1)",
            "  (1) EIRP ต่อแบนด์วิดท์อ้างอิง = pwr_ds_max + 10log₁₀(B_ref) + G      [B_ref: 4 kHz → +36.02 dB, 1 MHz → +60.00 dB]",
            "  (2) pfd(θ) = EIRP − 10log₁₀(4π·d(θ)²)      [การกระจายบนผิวทรงกลม free-space]",
            "  (3) เกณฑ์ T(θ) = P (0–5°) | P + r·(θ−5) (5–25°) | P + 20r (25–90°)      [NOTE 2 ของ Table 5-2; §1.1 คงที่ทุกมุม]",
            "  (4) NGSO: d(θ) = Rₑ·[√(((Rₑ+h)/Rₑ)² − cos²θ) − sinθ]      [h = perigee; Rₑ = 6 371 km]",
            "  (5) GSO: cosψ = cosφ·cosΔλ, d = √(Rₑ²+R²−2RₑR·cosψ), tanε = (cosψ−Rₑ/R)/sinψ      [R = 42 164 km]",
            "  (6) margin = pfd(θ*) − T(θ*)  โดย θ* คือมุมที่ margin มากที่สุด → margin > 0 = เกินเกณฑ์",
        ] + pfd_lines
    method_df = pd.DataFrame({"แสดงวิธีทำ 9.14 (ตรวจย้อนได้ทุกค่า)": pfd_lines}) if pfd_lines else pd.DataFrame()
    att6_df = pd.DataFrame(att6)
    if len(att6_df):
        att6_df["ข่ายงานดาวเทียม"] = att6_df["ข่ายงานดาวเทียม"].mask(
            att6_df["ข่ายงานดาวเทียม"].eq(att6_df["ข่ายงานดาวเทียม"].shift()), "")
    # ---- ผังการชนย่าน: เฉพาะมาตราที่ "เราเทียบเอง" คือ 9.21/A (เทียบ GSO ไทย)
    #      และ 9.21/B (เทียบ NGSO ไทย) — 9.7/9.12/9.12A เป็นผลคำนวณของ BR ไม่วาดทับ ----
    chart_src = {}
    for (prov, sec), bset in bands.items():
        if prov not in ("9.21/A", "9.21/B"):
            continue
        if not in_scope(sec, ids_a[(prov, sec)]):
            continue
        name, lon, adm = foreign_of(ids_a[(prov, sec)])
        chart_src[f"{name} — {prov}"] = {
            "adm": adm,
            "orbit": f"GSO {lon}" if lon.strip() not in ("", "?", "NGSO") else "NGSO",
            "bands": sorted((a_, b_, "".join(sorted(e_ - {""})))
                            for (a_, b_), e_ in bset.items()),
            "prov": prov, "sec": sec,
            "secn": int(sec_num(sec) or 999999),
            "scope": "gso" if prov == "9.21/A" else "ngso",
        }
    def _secsort(df, extra=None):
        if not len(df):
            return df
        k = df["Special section"].map(lambda x: int(sec_num(x) or 999999))
        d2 = df.assign(_k=k)
        by = (extra or []) + ["_k"]
        return d2.sort_values(by, kind="stable").drop(columns="_k").reset_index(drop=True)

    known = {"9.7", "9.12", "9.12A", "AP30"}
    other = {k: pd.DataFrame(v) for k, v in mem_b.items() if k not in known and v}
    return {
        "m97": _secsort(pd.DataFrame(mem_b["9.7"])),
        "m912": _secsort(pd.DataFrame(mem_b["9.12"])),
        "m912a": _secsort(pd.DataFrame(mem_b["9.12A"])),
        "ap30": _secsort(pd.DataFrame(mem_b["AP30"])),
        "m921": _secsort(pd.DataFrame(mem_a["9.21"]), extra=["มาตรา"]),
        "m914": _secsort(pd.DataFrame(mem_a["9.14"])),
        "flags": pd.DataFrame(flags_rows), "det": pd.DataFrame(det_a),
        "excluded": pd.DataFrame(excluded), "method": method_df, "att6": att6_df,
        "chart": chart_src, "other": other,
    }

# ================= UI =================

if not check_password():
    st.stop()

st.caption("โหมดเดียวจบ: วิเคราะห์ทั้ง **9.3 (API/A)** และ **CR/C (9.7 / 9.12 / 9.12A / 9.21 / 9.14)** จากไฟล์ชุดเดียว — portal CSV ใบเดียว (กด Reset all filters ก่อน Save list) มีครบทุก section อยู่แล้ว")

c1, c2 = st.columns(2)
mdb_file = c1.file_uploader("1) ไฟล์ ific.mdb", type=["mdb"])
portal_file = c2.file_uploader("2) portal CSV (รายการตีพิมพ์ทางการ ITU)", type=["csv"])

sheet_url = get_secret("thai_sheet_csv_url")
if sheet_url:
    st.success("ฐานความถี่ไทย: ใช้จาก **Google Sheet** ที่ตั้งไว้ใน Secrets "
               "(อัปโหลดไฟล์ด้านล่างเพื่อใช้แทนเฉพาะครั้งนี้ได้)")
else:
    st.info("ฐานความถี่ไทย: ยังไม่ได้ตั้ง Google Sheet — อัปโหลดไฟล์ด้านล่าง "
            'หรือเพิ่ม  thai_sheet_csv_url = "ลิงก์ CSV"  ใน Secrets')
thai_file = st.file_uploader("3) ฐานความถี่ไทย (Excel/CSV) — ข้ามได้ถ้าใช้ Google Sheet",
                             type=["xlsx", "csv"])

nfat_url = get_secret("nfat_sheet_csv_url")
if nfat_url:
    st.success("ตาราง NFAT (บริการภาคพื้นไทย): ใช้จาก **Google Sheet** ที่ตั้งไว้ใน Secrets "
               "(อัปโหลดไฟล์ด้านล่างเพื่อใช้แทนเฉพาะครั้งนี้ได้)")
else:
    st.info("ตาราง NFAT: ยังไม่ได้ตั้ง Google Sheet — อัปโหลดไฟล์ด้านล่าง "
            'หรือเพิ่ม  nfat_sheet_csv_url = "ลิงก์ CSV"  ใน Secrets '
            "(ใช้แสดงบริการภาคพื้นไทยในหน้า 9.14 และ 9.21/C)")
nfat_file = st.file_uploader("4) ตารางกำหนดคลื่นความถี่แห่งชาติ NFAT (Excel/CSV) — ข้ามได้ถ้าใช้ Google Sheet",
                             type=["xlsx", "csv"], key="nfat_up")
with st.expander("ℹ️ วิธีตั้ง Google Sheet เป็นฐานความถี่ไทย"):
    st.markdown(
        "1. ใน Google ชีต: **ไฟล์ → แชร์ → เผยแพร่ไปยังเว็บ**\n"
        "2. เลือกรูปแบบ **CSV** → เผยแพร่ → คัดลอกลิงก์ (ลงท้าย `output=csv`)\n"
        "3. ใส่ใน **Settings → Secrets**:  `thai_sheet_csv_url = \"ลิงก์ที่คัดลอก\"`\n"
        "4. ชีตต้องมีคอลัมน์: **network_name, freq_start_mhz, freq_end_mhz**"
    )
with st.expander("⚙️ ตัวกรองส่วน 9.3"):
    min_khz = st.number_input("กรองจุดทับซ้อนที่เล็กกว่า (kHz) — 0 = เอาทุกจุด (รวมแตะขอบ)",
                              min_value=0.0, value=0.0, step=1.0)

if st.button("▶ ประมวลผลทั้งฉบับ (9.3 + CR/C)", type="primary"):
    if not mdb_file:
        st.error("ต้องอัปโหลด ific.mdb"); st.stop()
    if not portal_file:
        st.error("ต้องอัปโหลด portal CSV (ขอบเขตทางการ) — กด Reset all filters แล้ว Save list as CSV"); st.stop()
    _orbit = st.empty()
    _orbit.markdown(
        """<div style="display:flex;align-items:center;gap:16px;padding:6px 2px">
<div style="position:relative;width:58px;height:58px;flex:none">
  <div style="position:absolute;inset:0;display:flex;align-items:center;justify-content:center;font-size:26px">🌏</div>
  <div style="position:absolute;inset:0;animation:ccorbit 2s linear infinite">
    <span style="position:absolute;top:-7px;left:50%;transform:translateX(-50%);font-size:17px">🛰️</span>
  </div>
</div>
<b style="color:#0E2A52">กำลังประมวลผล BR IFIC ทั้งฉบับ — อย่าเพิ่งปิดหน้า</b></div>
<style>@keyframes ccorbit{to{transform:rotate(360deg)}}</style>""",
        unsafe_allow_html=True)
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".mdb") as tf:
            tf.write(mdb_file.getvalue()); mdb_path = tf.name

        _stat = st.status("🛰️ กำลังประมวลผลทั้งฉบับ...", expanded=True)
        _stat.write("📂 อ่านไฟล์ MDB / ฐานความถี่ไทย / NFAT / portal CSV")
        ific = detect_ific(mdb_table(mdb_path, "notice"))
        try:
            thai, thai_src = load_thai(thai_file, sheet_url)
        except Exception as e:
            st.error(f"อ่านฐานความถี่ไทยไม่สำเร็จ: {e}"); st.stop()
        try:
            nfat, nfat_src = load_nfat(nfat_file, nfat_url)
        except Exception as e:
            nfat, nfat_src = None, None
            st.warning(f"อ่านตาราง NFAT ไม่สำเร็จ — ดำเนินการต่อโดยไม่มีคอลัมน์บริการภาคพื้น: {e}")

        prtl = read_portal_csv(portal_file)

        # ---------- ส่วน 9.3 (API/A) ----------
        _stat.write("📡 ส่วน 9.3 (API/A): เทียบความถี่กับฐานข่ายไทย")
        note93 = None
        foreign, apia, nodata = build_foreign_apia(mdb_path, prtl, ific)
        if thai is None:
            note93 = ("warning", "ไม่มีฐานความถี่ไทย — ข้ามการเทียบ 9.3 รอบนี้")
            detail = pd.DataFrame()
            memo93 = pd.DataFrame()
        else:
            if len(foreign):
                detail = overlap(foreign, thai, min_khz)
            else:
                detail = pd.DataFrame()
                note93 = ("info", "ไม่พบความถี่สำหรับ API/A รอบนี้")
            if len(detail):
                _ord = {r_.sat: int(sec_num(r_.section) or 999999)
                        for r_ in foreign.itertuples(index=False)}
                detail = (detail.assign(_k=detail["ดาวเทียมต่างชาติ"].map(_ord))
                          .sort_values(["_k", "ตปท_f_min", "ข่ายไทย"], kind="stable")
                          .drop(columns="_k").reset_index(drop=True))
            if len(detail) == 0:
                detail = pd.DataFrame(columns=["ดาวเทียมต่างชาติ", "ประเภท", "ข่ายไทย"])
                if note93 is None:
                    note93 = ("success", "ไม่มีความถี่ API/A ใดทับซ้อนกับฐานไทย")
            memo93 = memo_93(detail, foreign, nodata) if (len(detail) or nodata) else pd.DataFrame()

        chart93 = {}
        if len(foreign) and thai is not None:
            for sat, g in foreign.groupby("sat"):
                ivd = {}
                for r in g.itertuples(index=False):
                    ivd.setdefault((r.f_min, r.f_max), set()).add(r.emi)
                m0 = g.iloc[0]
                chart93[f"{sat} — 9.3"] = {
                    "adm": m0.adm, "orbit": m0.orbit,
                    "bands": sorted((a, b, "".join(sorted(e - {""})))
                                    for (a, b), e in ivd.items()),
                    "prov": "9.3", "sec": m0.section,
                    "secn": int(sec_num(m0.section) or 999999),
                    "scope": "all",
                }

        # ---------- ส่วน CR/C ----------
        _stat.write("🛰️ ส่วน CR/C: 9.7 / 9.12 / 9.12A / 9.21 / 9.14 (คำนวณ pfd)")
        crc_portal = [r for r in prtl if col(r, "Special section").startswith("CR")]
        allowed_secnums = {col(r, "Special section number").strip() for r in crc_portal}
        allowed_ids = set()
        for r in crc_portal:
            for x in (col(r, "Notice ID"), col(r, "Targeted notice ID")):
                if x:
                    allowed_ids.add(x.strip())
        if thai is None:
            st.warning("ยังไม่มีฐานความถี่ไทย — 9.21 จะไม่มีคอลัมน์ข่ายไทยที่ทับ "
                       "(9.14 ไม่ใช้ฐานไทย — คัดกรองด้วย pfd)")

        R = build_crc(mdb_path, ific, thai, allowed_secnums, allowed_ids, nfat)
        os.unlink(mdb_path)

        used_nums = set()
        for df in (R["m97"], R["m912"], R["m912a"], R["m921"], R["m914"]):
            if len(df):
                used_nums |= {sec_num(s) for s in df["Special section"]}
        no_tha = [f"{r['Special section']}/{r['Special section number']} {r['Satellite name']}"
                  for r in crc_portal
                  if col(r, "Special section number").strip() not in used_nums]

        _stat.write("📦 สรุปรวม + สร้าง Excel + ร่างบันทึก Word")
        summary = build_summary(prtl, memo93, R)
        charts = {**chart93, **R["chart"]}

        W7X = [20, 12, 8, 18, 8, 60, 10]
        wb = [("สรุปรวม", summary, [7, 24, 20, 9] + [7] * len(SUM_PROVS) + [18]),
              ("9.3 สำหรับบันทึก", memo93, [20, 12, 18, 8, 60, 30, 9, 8]),
              ("9.3 รายละเอียด", detail, [20, 8, 12, 14, 14, 16, 12, 12, 24, 12, 12, 12]),
              ("9.7", R["m97"], W7X), ("9.12", R["m912"], W7X), ("9.12A", R["m912a"], W7X),
              ("AP30-30A", R["ap30"], W7X),
              ("9.21", R["m921"], [20, 12, 7, 16, 7, 30, 40, 30, 30, 24, 9]),
              ("9.14", R["m914"], [20, 16, 7, 14, 7, 18, 13, 22, 20, 9, 30, 30, 7]),
              ("เอกสารแนบ6 9.14", R["att6"], [20, 14, 13, 17, 22, 14]),
              ("9.14 วิธีทำ", R["method"], [120]),
              ("นอกขอบเขต", R["excluded"], [10, 18, 22, 44]),
              ("flags 9.7-9.12A", R["flags"], [20, 10, 18, 22, 10, 10]),
              ("รายละเอียดย่าน 9.21", R["det"], [20, 10, 18, 12, 12, 10])]
        for k, dfo in R["other"].items():
            wb.append((f"มาตรา {k}", dfo, W7X))
        xls = build_excel([(t, d, w) for t, d, w in wb if d is not None and len(d)],
                          green_sheets={"สรุปรวม"})

        docx_bytes, docx_err = None, None
        if _HAS_DOCX:
            try:
                docx_bytes = build_memo_docx(ific, R["m97"], R["m912"], R["m912a"],
                                             R["m921"], R["m914"], memo93)
            except Exception as e:
                docx_err = f"สร้างร่างบันทึกไม่สำเร็จ: {e}"
        else:
            docx_err = "ยังไม่ได้เพิ่ม python-docx ใน requirements.txt"

        st.session_state["RESA"] = {
            "ific": ific, "thai": thai, "thai_src": thai_src,
            "thai_len": len(thai) if thai is not None else 0,
            "nfat_src": nfat_src, "nfat_len": len(nfat) if nfat is not None else 0,
            "apia": len(apia), "nodata": nodata, "note93": note93,
            "nsat": foreign["sat"].nunique() if len(foreign) else 0,
            "detail": detail, "memo93": memo93,
            "R": R, "no_tha": no_tha, "ncrc": len(crc_portal),
            "summary": summary, "charts": charts,
            "xls": xls, "docx": docx_bytes, "docx_err": docx_err,
        }
        st.session_state.pop("charts_pdf", None)
        _stat.update(label="✅ ประมวลผลเสร็จ — เลื่อนลงดูผล 4 แท็บด้านล่าง",
                     state="complete", expanded=False)
        _orbit.empty()
    except Exception as e:
        _orbit.empty()
        st.session_state.pop("RESA", None)
        st.error(f"เกิดข้อผิดพลาด: {e}")

# ================= แสดงผลล่าสุด (คงอยู่แม้กดดาวน์โหลด/สลับแท็บ) =================
RESA = st.session_state.get("RESA")
if RESA:
    ific = RESA["ific"]
    R = RESA["R"]
    tb_sum, tb_93, tb_crc, tb_chart = st.tabs(
        ["📌 สรุปรวม", "9.3 API/A", "CR/C 9.7–9.14", "📊 ผังชนย่าน"])

    # ---------- แท็บสรุปรวม ----------
    with tb_sum:
        s = RESA["summary"]
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("IFIC", ific)
        m2.metric("ตีพิมพ์ทั้งหมด", len(s))
        m3.metric("เกี่ยวข้องกับไทย", int((s[SUM_PROVS] != "").any(axis=1).sum()) if len(s) else 0)
        m4.metric("9.3", int((s["9.3"] != "").sum()) if len(s) else 0)
        m5.metric("CR/C", int((s[[p for p in SUM_PROVS if p != "9.3"]] != "")
                              .any(axis=1).sum()) if len(s) else 0)
        st.caption("X = กระทบ/เข้าเกณฑ์ · x = แตะขอบ (0 MHz) · Remark เว้นว่างให้กรอกเอง "
                   "(เช่น หนังสือบริษัทที่ร้องขอ)")
        if len(s):
            try:
                sty = s.style.map(
                    lambda v: ("background-color:#92D050;font-weight:bold;text-align:center"
                               if v == "X" else
                               "background-color:#E2EFDA;text-align:center" if v == "x" else ""),
                    subset=SUM_PROVS)
                st.dataframe(sty, width="stretch", height=min(38 * len(s) + 40, 700))
            except Exception:
                st.dataframe(s, width="stretch")
        else:
            st.info("ไม่มีรายการตีพิมพ์ใน portal CSV")
        st.download_button("⬇ ดาวน์โหลด Excel (ทุกชีต รวม 9.3 + CR/C)", RESA["xls"],
                           file_name=f"IFIC{ific}_ผลตรวจทั้งฉบับ.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        if RESA.get("docx"):
            st.download_button("📄 ดาวน์โหลดร่างบันทึกข้อความ (.docx)", RESA["docx"],
                               file_name=f"ร่างบันทึก_IFIC{ific}.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            st.caption("ร่างตามฟอร์ม ดบ. — รวมทุกมาตรา (9.7 · 9.12+9.12A · 9.21 · 9.14 · 9.3) "
                       "ช่องที่แอปไม่มีทางรู้ใส่ (ระบุ) ไว้ให้เติม")
        elif RESA.get("docx_err"):
            st.info(f"ร่างบันทึก Word: {RESA['docx_err']}")

    # ---------- แท็บ 9.3 ----------
    with tb_93:
        st.caption(f"ฐานความถี่ไทย: {RESA['thai_src']} ({RESA['thai_len']} ย่าน)")
        m1, m2, m3 = st.columns(3)
        m1.metric("API/A (ขอบเขต)", RESA["apia"])
        m2.metric("มีความถี่", RESA["nsat"])
        m3.metric("ไม่มี tech data", len(RESA["nodata"]))
        if RESA["nodata"]:
            st.warning("ไม่มี technical data (publication-only): "
                       + ", ".join(x[0] for x in RESA["nodata"]))
        if RESA["note93"]:
            getattr(st, RESA["note93"][0])(RESA["note93"][1])
        detail, memo93 = RESA["detail"], RESA["memo93"]
        if len(detail):
            n_real = int((detail["ประเภท"] == "ทับจริง").sum())
            n_edge = int((detail["ประเภท"] == "แตะขอบ (0 MHz)").sum())
            st.caption(f"จุดทับซ้อนทั้งหมด {len(detail)} — ทับจริง {n_real} | แตะขอบ (0 MHz) {n_edge}")
        if memo93 is not None and len(memo93):
            st.subheader("📋 ตารางสำหรับบันทึก (1 แถว/ข่าย)")
            st.dataframe(memo93, width="stretch")
        if len(detail):
            st.subheader("รายละเอียดจุดทับซ้อน")
            st.dataframe(detail, width="stretch")

    # ---------- แท็บ CR/C ----------
    with tb_crc:
        if RESA["nfat_len"]:
            st.caption(f"ตาราง NFAT: {RESA['nfat_src']} ({RESA['nfat_len']} แถวจัดสรร)")
        st.caption(f"ขอบเขต: รายการ CR/C ทางการจาก portal {RESA['ncrc']} รายการ")
        if len(R["excluded"]):
            st.warning("ข่ายที่พบใน MDB แต่**ไม่อยู่ในรายการตีพิมพ์รอบนี้** — ตัดออกจากผล "
                       "(ตรวจสอบได้ในชีต 'นอกขอบเขต'):")
            st.dataframe(R["excluded"], width="stretch")
        if RESA["no_tha"]:
            st.caption("ตีพิมพ์รอบนี้แต่ไม่ระบุ THA (ไม่ต้องดำเนินการ): " + ", ".join(RESA["no_tha"]))

        m914 = R["m914"]
        n914 = m914["ข่ายงานดาวเทียม"].nunique() if len(m914) else 0
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("9.7", len(R["m97"]))
        m2.metric("9.12", len(R["m912"]))
        m3.metric("9.12A", len(R["m912a"]))
        m4.metric("9.21/A-C", len(R["m921"]))
        m5.metric("9.14", n914)

        for title, df in [("มาตรา 9.7 — BR ระบุชื่อข่ายไทย (ผลคำนวณของ BR)", R["m97"]),
                          ("มาตรา 9.12 — BR ระบุชื่อข่ายไทย", R["m912"]),
                          ("มาตรา 9.12A — BR ระบุชื่อข่ายไทย", R["m912a"]),
                          ("มาตรา 9.21 — /A เทียบข่าย GSO ไทย · /B เทียบข่าย NGSO ไทย · /C สถานีภาคพื้น (ดูคอลัมน์ NFAT)", R["m921"])]:
            st.subheader(title)
            if len(df):
                st.dataframe(df, width="stretch")
            else:
                st.info("ไม่มีในรอบนี้")
        if len(R["ap30"]):
            st.subheader("AP30/30A Article 7 — BR ระบุชื่อข่ายไทย")
            st.dataframe(R["ap30"], width="stretch")
        for k, dfo in R["other"].items():
            st.subheader(f"มาตรา {k} — BR ระบุชื่อข่ายไทย")
            st.dataframe(dfo, width="stretch")

        st.subheader("มาตรา 9.14 — คัดกรอง pfd เทียบเกณฑ์ AP5 (คุ้มครองบริการภาคพื้นของไทย)")
        if len(m914):
            n_over = int(m914["ผลการพิจารณา"].astype(str).str.startswith("เกิน").sum())
            n_foot = int(m914["ผลการพิจารณา"].astype(str).str.contains("footnote").sum())
            st.caption(f"ข่าย {n914} | แถวเทียบ {len(m914)} — เกินเกณฑ์ {n_over} | ต้องตรวจ footnote {n_foot}")
            st.dataframe(m914, width="stretch")
            st.markdown("**ตารางสรุปแบบเอกสารแนบ (ต่อข่าย — ก๊อปลง Word ได้)**")
            st.dataframe(R["att6"], width="stretch")
            with st.expander("📐 แสดงวิธีทำ 9.14 — สูตร + แทนค่าทุกขั้น (แนบผู้บริหาร/ตรวจย้อนได้)"):
                mm = R["method"]
                st.text("\n".join(mm[mm.columns[0]].tolist()) if len(mm) else "")
        else:
            st.info("ไม่มีในรอบนี้")

    # ---------- แท็บผังชนย่าน ----------
    with tb_chart:
        st.caption("วาดเฉพาะมาตราที่ **ปบ. เทียบเอง**: 9.3 (เทียบฐานข่ายไทยทั้งหมด), "
                   "9.21/A (เทียบ GSO ไทย), 9.21/B (เทียบ NGSO ไทย) — "
                   "ส่วน 9.7/9.12/9.12A เป็นผลคำนวณของ BR จึงไม่วาดทับ")
        charts = RESA["charts"]
        if not _HAS_MPL:
            st.info("ยังไม่ได้เพิ่ม matplotlib ใน requirements.txt — เพิ่มแล้วแอปจะวาดผังให้อัตโนมัติ")
        elif RESA["thai"] is None:
            st.info("ต้องมีฐานความถี่ไทยจึงจะวาดผังได้")
        elif not charts:
            st.info("รอบนี้ไม่มีรายการ 9.3 / 9.21A / 9.21B ที่ทับกับฐานไทยให้วาด")
        else:
            if not _THAI_FONT_OK:
                st.warning("ไม่พบไฟล์ Sarabun-Regular.ttf / Sarabun-Bold.ttf ใน repo — "
                           "ตัวหนังสือไทยในผังจะแสดงไม่ได้จนกว่าจะอัปโหลดฟอนต์ (วางที่ root ของ repo ได้เลย)")
            opts = sorted(charts, key=lambda k: (charts[k]["prov"] != "9.3",
                                                 charts[k].get("secn", 999999), k))
            sel = st.selectbox("เลือกรายการ (ข่าย — มาตรา)", opts, key="cc_sel")
            fig, why = collision_fig(sel, charts[sel], RESA["thai"], ific)
            if fig is None:
                st.info(why)
            else:
                st.pyplot(fig)
                _buf = io.BytesIO()
                fig.savefig(_buf, format="png", dpi=150, bbox_inches="tight", facecolor="white")
                plt.close(fig)
                st.download_button("⬇ ดาวน์โหลดผังเป็น PNG", _buf.getvalue(),
                                   file_name=f"ผังชนย่าน_{sel.replace(' — ', '_').replace('/', '-')}_IFIC{ific}.png",
                                   mime="image/png", key="cc_png")
            st.divider()
            st.markdown("**📚 PDF รวมเล่มทุกรายการ** — หน้าแรกเป็นสารบัญครบทุกรายการที่ตรวจ "
                        "(รายการที่ไม่พบการทับระบุผลไว้ชัด) ตามด้วยผังหน้าละรายการ")
            if st.button(f"🗜 สร้าง PDF รวมเล่ม ({len(charts)} รายการ)", key="pdf_all_btn"):
                _pg = st.progress(0.0, text="🛰️ กำลังวาดผังทีละรายการ...")
                try:
                    st.session_state["charts_pdf"] = charts_pdf(
                        charts, RESA["thai"], ific,
                        progress=lambda f: _pg.progress(f, text=f"🛰️ กำลังวาดผัง... {int(f*100)}%"))
                    _pg.progress(1.0, text="เสร็จแล้ว")
                except Exception as e:
                    st.error(f"สร้าง PDF ไม่สำเร็จ: {e}")
            if st.session_state.get("charts_pdf"):
                st.download_button("⬇ ดาวน์โหลด PDF รวมเล่ม",
                                   st.session_state["charts_pdf"],
                                   file_name=f"ผังชนย่านทั้งหมด_IFIC{ific}.pdf",
                                   mime="application/pdf", key="pdf_all_dl")
