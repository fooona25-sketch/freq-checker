"""
Core engine for the Thai satellite frequency cross-check tool.

Design principle: this module does the DETERMINISTIC work only
(read BR IFIC database -> extract foreign frequencies -> compare with the
Thai reference). No AI/LLM is involved here, so results are exact and
reproducible. The optional letter-drafting step (phase 2) lives elsewhere.
"""

import io
import subprocess
import pandas as pd

# Special-section provision codes as stored in the SNS "prov" field.
PROV_BY_SECTION = {
    "API/A": "9.1/IA",   # advance publication -> target of RR No. 9.3 comments
    "CR/C":  "9.6",      # coordination request
    "Notif": "11.2",     # notification under Article 11
}

# emi_rcp values in the BR grp table
DIRECTION_LABEL = {"E": "Tx (E)", "R": "Rx (R)"}


def _mdb_table(mdb_path: str, table: str) -> pd.DataFrame:
    """Export a single table from an Access .mdb to a DataFrame via mdbtools.
    Only the tables we actually need are read, so memory stays low even on
    a 100+ MB IFIC file."""
    out = subprocess.run(
        ["mdb-export", mdb_path, table],
        capture_output=True, text=True, check=True,
    ).stdout
    return pd.read_csv(io.StringIO(out), dtype=str, keep_default_na=False)


def extract_foreign(mdb_path: str, section: str = "API/A") -> pd.DataFrame:
    """Return one row per foreign frequency group for the chosen special
    section. Frequencies are exact numeric values straight from the database
    (no OCR)."""
    prov = PROV_BY_SECTION.get(section, "9.1/IA")
    com = _mdb_table(mdb_path, "com_el")
    grp = _mdb_table(mdb_path, "grp")

    nets = com[com["prov"] == prov][
        ["ntc_id", "adm", "sat_name", "long_nom"]
    ].drop_duplicates("ntc_id")

    g = grp[grp["ntc_id"].isin(nets["ntc_id"]) & (grp["freq_min"] != "")].copy()
    g = g.merge(nets, on="ntc_id", how="left")
    g["freq_min"] = g["freq_min"].astype(float)
    g["freq_max"] = g["freq_max"].astype(float)
    g["type"] = g["long_nom"].apply(
        lambda x: "GSO" if str(x).strip() not in ("", "nan") else "NGSO"
    )
    g["direction"] = g["emi_rcp"].map(DIRECTION_LABEL).fillna(g["emi_rcp"])

    out = g[[
        "adm", "sat_name", "type", "beam_name", "direction",
        "emi_rcp", "freq_min", "freq_max", "long_nom",
    ]].rename(columns={
        "sat_name": "satellite", "beam_name": "beam",
        "freq_min": "f_min_mhz", "freq_max": "f_max_mhz",
        "long_nom": "orbital_long",
    })
    return out.reset_index(drop=True)


def normalise_thai(df: pd.DataFrame) -> pd.DataFrame:
    """Accept a Thai reference table with flexible column names and return a
    standard shape. Required: a name + a start + an end frequency in MHz.
    Optional: direction (E/R or Tx/Rx) and type (GSO/NGSO)."""
    cols = {c.lower().strip(): c for c in df.columns}

    def pick(*cands):
        for c in cands:
            if c in cols:
                return cols[c]
        return None

    name = pick("network_name", "network", "name", "ชื่อข่าย")
    fmin = pick("freq_start_mhz", "freq_min", "f_min_mhz", "freq_start", "ความถี่ต่ำ")
    fmax = pick("freq_end_mhz", "freq_max", "f_max_mhz", "freq_end", "ความถี่สูง")
    direc = pick("direction", "dir", "emi_rcp", "ทิศทาง")
    typ = pick("type", "gso_ngso", "ประเภท")
    if not (name and fmin and fmax):
        raise ValueError(
            "Thai reference must have a network name + start/end frequency (MHz). "
            f"Found columns: {list(df.columns)}"
        )

    out = pd.DataFrame({
        "thai_network": df[name].astype(str),
        "thai_min": pd.to_numeric(df[fmin], errors="coerce"),
        "thai_max": pd.to_numeric(df[fmax], errors="coerce"),
    })
    out["thai_dir"] = df[direc].astype(str) if direc else ""
    out["thai_type"] = df[typ].astype(str) if typ else ""
    out = out.dropna(subset=["thai_min", "thai_max"])
    return out.reset_index(drop=True)


def compute_overlaps(foreign: pd.DataFrame, thai: pd.DataFrame,
                     min_overlap_khz: float = 0.0) -> pd.DataFrame:
    """Deterministic band-overlap test between every foreign group and every
    Thai band. Two bands overlap when max(lows) < min(highs).
    min_overlap_khz lets the user ignore negligible slivers."""
    fa = foreign.to_dict("records")
    ta = thai.to_dict("records")
    thr = min_overlap_khz / 1000.0  # to MHz
    rows = []
    for f in fa:
        for t in ta:
            lo = max(f["f_min_mhz"], t["thai_min"])
            hi = min(f["f_max_mhz"], t["thai_max"])
            w = hi - lo
            if w > thr:
                rows.append({
                    "foreign_adm": f["adm"],
                    "foreign_satellite": f["satellite"],
                    "foreign_type": f["type"],
                    "foreign_beam": f["beam"],
                    "foreign_dir": f["direction"],
                    "foreign_f_min": round(f["f_min_mhz"], 4),
                    "foreign_f_max": round(f["f_max_mhz"], 4),
                    "thai_network": t["thai_network"],
                    "thai_dir": t.get("thai_dir", ""),
                    "thai_f_min": round(t["thai_min"], 4),
                    "thai_f_max": round(t["thai_max"], 4),
                    "overlap_min": round(lo, 4),
                    "overlap_max": round(hi, 4),
                    "overlap_mhz": round(w, 4),
                })
    return pd.DataFrame(rows)


def summarise(matches: pd.DataFrame) -> pd.DataFrame:
    """One row per foreign network that overlaps any Thai network."""
    if matches.empty:
        return matches
    return (matches.groupby(["foreign_adm", "foreign_satellite", "foreign_type"])
            .agg(overlaps=("thai_network", "size"),
                 thai_networks=("thai_network", lambda x: ", ".join(sorted(set(x)))))
            .reset_index().sort_values("overlaps", ascending=False))


def to_excel_bytes(summary: pd.DataFrame, detail: pd.DataFrame) -> bytes:
    """Build the downloadable Excel (summary + detail) in memory."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xl:
        summary.to_excel(xl, sheet_name="สรุป", index=False)
        detail.to_excel(xl, sheet_name="รายละเอียด", index=False)
    return buf.getvalue()
