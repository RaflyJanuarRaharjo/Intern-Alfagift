"""Report Daily MDS (Mini Darkstore).

Upload Detail_Data_*.csv (+ opsional Master MDS) -> Report Daily MDS.xlsx.
Logika dipindah apa adanya dari Report_Daily_MDS.py (halaman Streamlit lama).
"""
from __future__ import annotations

import io

import os

import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

DEFAULT_MASTER = [
    ("1GQ7", "CHRISTOPEL MIHING"),
    ("1GF4", "CILIK RIWUT KM 1"),
    ("1GQ3", "DIPANDJAITAN SAMPIT"),
    ("1G5D", "KARET"),
    ("1GE4", "PULANG PISAU"),
    ("1GF5", "TEMANGGUNG TILUNG 2"),
    ("X812", "SUSUKAN"),
    ("CG13", "CILINCING KELAPA"),
    ("1J10", "PERMATA SARI INDAH"),
    ("K471", "CITRA"),
    ("TF45", "GOLF ISLAND"),
    ("KA77", "PURI MANSION"),
    ("TD40", "SUVARNA TERRACE 8"),
    ("X528", "CIBINONG GN SINDUR"),
    ("P304", "TUGU HARUM"),
    ("J566", "TIPAR CAKUNG 1 - 3"),
    ("R236", "AHMAD YANI POLMAN"),
    ("R202", "H.A.MUH. ARSYAD 2"),
    ("R151", "MACCOPA MAROS"),
    ("R184", "MH.THAMRIN BONE"),
    ("R264", "TAMANGAPA RAYA 3"),
    ("R287", "SLTN HSNDDIN PINRANG"),
    ("1D73", "LEBAY HASAN"),
    ("1D74", "PROF. SRI SOEDEWI"),
    ("1P35", "KRIDASANA"),
    ("1P41", "KOMYOS SUDARSO"),
    ("1G65", "KARANG ANYAR"),
    ("W227", "KUTACANE KABANJAHE"),
    ("1D85", "PERINTIS"),
    ("R385", "ANDI RADJA BULUKUMBA"),
    ("1PG2", "ADI SUCIPTO KM 8"),
    ("L510", "SUMBER REJO 2"),
    ("1GM1", "BOEJASIN 2"),
    ("1A08", "PUSKOPKAR"),
    ("L548", "CITRA GARDEN 2"),
    ("1GC1", "PALAM RAYA"),
    ("Q389", "BAKTI SERAGA"),
    ("1AJ4", "AHMAD YANI- DUMAI"),
    ("2D07", "MERAPI SUBUR"),
    ("R446", "MARTADINATA MAMUJU"),
    ("1PH9", "KOTA BARU"),
    ("1GD4", "RAYA BATULICIN"),
    ("1P19", "JEMBATAN KAPUAS II"),
    ("1GC2", "AHMAD YANI KM 12"),
    ("1A0Y", "JENSUDPULAUBENGKALIS"),
    ("UA24", "GEDUNGAN SUMENEP"),
    ("1Y02", "LANSOT"),
    ("2DO6", "SIMP BUKIT LESTARI"),
    ("1DR8", "JENSUD 2 TUNGKAL"),
    ("UD50", "SEMEMI SBY"),
    ("2DR5", "MCDERMOTT"),
    ("JD65", "CITRA INDAH AGAVE"),
]

BULAN = ["", "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli",
         "Agustus", "September", "Oktober", "November", "Desember"]
BULAN_SHORT = ["", "Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul",
               "Agu", "Sept", "Okt", "Nov", "Des"]

AGG_COLS = ["KD_STORE", "NAMA_STORE", "NAMA_BRANCH", "Sum of JHK", "Sum of SALES",
            "Sum of SALES_TAGI", "Average of SPD", "Average of STD",
            "Average of APC", "Average of PERCENT_GM"]


# ----------------------------------------------------------------------------
# Load
# ----------------------------------------------------------------------------
def _read_any(src, **kw) -> pd.DataFrame:
    """Baca CSV/XLSX dari path atau file-like (mis. Streamlit UploadedFile)."""
    name = str(getattr(src, "name", src)).lower()
    if name.endswith((".xlsx", ".xls")):
        return pd.read_excel(src, **kw)
    return pd.read_csv(src, **kw)


def load_raw(src) -> pd.DataFrame:
    df = _read_any(src, dtype={"KD_STORE": str, "KD_BRANCH": str})
    df.columns = [c.strip().upper() for c in df.columns]
    df["TANGGAL"] = pd.to_datetime(df["TANGGAL"]).dt.normalize()
    df["KD_STORE"] = df["KD_STORE"].str.strip().str.upper()
    return df


def load_master(src=None) -> pd.DataFrame:
    if src is None:
        return pd.DataFrame(DEFAULT_MASTER, columns=["KD_STORE", "NAMA_STORE_MASTER"])
    m = _read_any(src, dtype=str)
    m.columns = [c.strip().upper().replace("_", " ") for c in m.columns]
    m = m.rename(columns={"KD STORE": "KD_STORE", "NAMA STORE": "NAMA_STORE_MASTER"})
    m["KD_STORE"] = m["KD_STORE"].str.strip().str.upper()
    return m.drop_duplicates("KD_STORE")


# ----------------------------------------------------------------------------
# Transform
# ----------------------------------------------------------------------------
def _agg(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby(["KD_STORE", "NAMA_STORE", "NAMA_BRANCH"], as_index=False).agg(
        **{
            "Sum of JHK": ("JHK", "sum"),
            "Sum of SALES": ("SALES", "sum"),
            "Sum of SALES_TAGI": ("SALES_TAGI", "sum"),
            "Average of SPD": ("SPD", "mean"),
            "Average of STD": ("STD", "mean"),
            "Average of APC": ("APC", "mean"),
            "Average of PERCENT_GM": ("PERCENT_GM", "mean"),
        }
    )
    return g[AGG_COLS]


def build(raw: pd.DataFrame, master: pd.DataFrame, tanggal=None) -> dict:
    tgl = pd.Timestamp(tanggal).normalize() if tanggal else raw["TANGGAL"].max()
    awal = tgl.replace(day=1)

    data_mds = raw[raw["KD_STORE"].isin(master["KD_STORE"])].copy()
    data_mds.insert(data_mds.columns.get_loc("REMARK") + 1, "REMARK MDS", "MINI DARKSTORE")

    mtd_rows = data_mds[(data_mds["TANGGAL"] >= awal) & (data_mds["TANGGAL"] <= tgl)]
    day_rows = data_mds[data_mds["TANGGAL"] == tgl]
    if day_rows.empty:
        raise ValueError(f"Tidak ada data MDS untuk tanggal {tgl:%Y-%m-%d}")

    daily, mtd = _agg(day_rows), _agg(mtd_rows)

    s = mtd.merge(daily, on=["KD_STORE"], how="left", suffixes=("_mtd", "_d"))
    summary = pd.DataFrame({
        "Kode Toko": s["KD_STORE"],
        "Nama Toko": s["NAMA_STORE_mtd"],
        "Cabang": s["NAMA_BRANCH_mtd"],
        "D_JHK": s["Sum of JHK_d"],
        "D_SPD": s["Average of SPD_d"],
        "D_SPD_TAGI": s["Sum of SALES_TAGI_d"],
        "D_STD": s["Average of STD_d"],
        "D_APC": s["Average of APC_d"],
        "D_GM": s["Average of PERCENT_GM_d"],
        "M_JHK": s["Sum of JHK_mtd"],
        "M_SALES": s["Sum of SALES_mtd"],
        "M_SALES_TAGI": s["Sum of SALES_TAGI_mtd"],
        "M_SPD": s["Average of SPD_mtd"],
        "M_SPD_TAGI": s["Sum of SALES_TAGI_mtd"] / s["Sum of JHK_mtd"],
        "M_STD": s["Average of STD_mtd"],
        "M_APC": s["Average of APC_mtd"],
        "M_GM": s["Average of PERCENT_GM_mtd"],
    }).sort_values("M_SALES", ascending=False, ignore_index=True)

    missing = sorted(set(master["KD_STORE"]) - set(data_mds["KD_STORE"]))
    return dict(tgl=tgl, awal=awal, summary=summary, daily=daily, mtd=mtd,
                data_mds=data_mds.sort_values(["TANGGAL", "KD_STORE"]), missing=missing)


# ----------------------------------------------------------------------------
# Write Excel
# ----------------------------------------------------------------------------
FILL = PatternFill("solid", fgColor="DDEBF7")
THIN = Side(style="thin")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
BOLD = Font(bold=True)


def _write_table(ws, df: pd.DataFrame, num_fmt: dict | None = None):
    ws.append(list(df.columns))
    for c in ws[1]:
        c.font, c.fill = BOLD, FILL
    for row in df.itertuples(index=False):
        ws.append([None if pd.isna(v) else (v.to_pydatetime() if isinstance(v, pd.Timestamp) else v)
                   for v in row])
    for i, col in enumerate(df.columns, 1):
        ws.column_dimensions[get_column_letter(i)].width = max(10, min(28, len(str(col)) + 2))
        if num_fmt and col in num_fmt:
            for (cell,) in ws.iter_rows(min_row=2, min_col=i, max_col=i):
                cell.number_format = num_fmt[col]
    ws.freeze_panes = "A2"


def to_excel(res: dict) -> bytes:
    tgl, awal, sm = res["tgl"], res["awal"], res["summary"]
    bln, thn = BULAN[tgl.month], tgl.year
    d_label = f"{tgl.day} {bln} {thn}"
    p_label = f"{awal.day} - {tgl.day} {bln} {thn}"
    sh_day = f"Data {tgl.day} {BULAN_SHORT[tgl.month]} {thn}"
    sh_mtd = f"Data MTD {awal.day} - {tgl.day} {BULAN_SHORT[tgl.month]} {thn}"

    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws.sheet_view.showGridLines = False

    ws["B2"] = "Daily Report Performance Mini Darkstore"
    ws["B3"] = f"Periode {p_label}"
    ws["B2"].font = ws["B3"].font = BOLD

    # Header
    for col, txt in zip("BCDE", ["No", "Kode Toko", "Nama Toko", "Cabang"]):
        ws[f"{col}5"] = txt
        ws.merge_cells(f"{col}5:{col}6")
    ws["F5"] = f"Sales Mini Darkstore {d_label}"
    ws.merge_cells("F5:K5")
    ws["L5"] = f"Sales Mini Darkstore {p_label}"
    ws.merge_cells("L5:S5")
    sub = ["JHK", "SPD", "SPD Tag I", "STD", "APC", "%GM",
           "JHK", "Total Net Sales", "Total Net Sales Tag I", "SPD", "SPD Tag I", "STD", "APC", "%GM"]
    for i, txt in enumerate(sub):
        ws.cell(6, 6 + i, txt)
    for r in (5, 6):
        for c in range(2, 20):
            cell = ws.cell(r, c)
            cell.font, cell.fill, cell.alignment, cell.border = BOLD, FILL, CENTER, BORDER

    # Body
    keys = ["D_JHK", "D_SPD", "D_SPD_TAGI", "D_STD", "D_APC", "D_GM",
            "M_JHK", "M_SALES", "M_SALES_TAGI", "M_SPD", "M_SPD_TAGI", "M_STD", "M_APC", "M_GM"]
    fmts = ["0", "#,##0", "#,##0", "#,##0", "#,##0", "0.00%",
            "0", "#,##0", "#,##0", "#,##0", "#,##0", "#,##0", "#,##0", "0.00%"]
    first = 7
    for n, row in sm.iterrows():
        r = first + n
        vals = [n + 1, row["Kode Toko"], row["Nama Toko"], row["Cabang"]] + [row[k] for k in keys]
        for c, v in enumerate(vals, 2):
            cell = ws.cell(r, c, None if pd.isna(v) else v)
            cell.border, cell.alignment = BORDER, Alignment(horizontal="left" if c == 4 else "center")
            if c >= 6:
                cell.number_format = fmts[c - 6]
    last = first + len(sm) - 1

    # Total row
    tr = last + 1
    ws.cell(tr, 2, "Average/Sum")
    ws.merge_cells(start_row=tr, start_column=2, end_row=tr, end_column=5)
    for c in range(6, 20):
        L = get_column_letter(c)
        fn = "SUM" if L in ("M", "N") else "AVERAGE"
        ws.cell(tr, c, f"={fn}({L}{first}:{L}{last})").number_format = fmts[c - 6]
    for c in range(2, 20):
        cell = ws.cell(tr, c)
        cell.font, cell.fill, cell.alignment, cell.border = BOLD, FILL, CENTER, BORDER

    ws.conditional_formatting.add(
        f"M{first}:M{last}",
        ColorScaleRule(start_type="min", start_color="FFEF9C", end_type="max", end_color="63BE7B"))

    widths = dict(A=2, B=5, C=10, D=26, E=14, F=6, G=13, H=13, I=7, J=10, K=8,
                  L=6, M=16, N=18, O=13, P=12, Q=7, R=10, S=8)
    for k, v in widths.items():
        ws.column_dimensions[k].width = v
    ws.row_dimensions[5].height = 30
    ws.freeze_panes = "F7"

    # Supporting sheets
    pfmt = {"Sum of SALES": "#,##0", "Sum of SALES_TAGI": "#,##0", "Average of SPD": "#,##0",
            "Average of STD": "#,##0.0", "Average of APC": "#,##0", "Average of PERCENT_GM": "0.00%"}
    _write_table(wb.create_sheet(sh_day[:31]), res["daily"], pfmt)
    _write_table(wb.create_sheet(sh_mtd[:31]), res["mtd"], pfmt)
    _write_table(wb.create_sheet("Data MDS"), res["data_mds"], {"TANGGAL": "yyyy-mm-dd"})

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def generate(raw_src, master_src=None, tanggal=None) -> tuple[bytes, dict]:
    res = build(load_raw(raw_src), load_master(master_src), tanggal)
    return to_excel(res), res


def output_name(res: dict) -> str:
    t = res["tgl"]
    return f"Report Daily MDS {t.day} {BULAN[t.month]} {t.year}.xlsx"



def available_dates(raw_path):
    return sorted(load_raw(raw_path)["TANGGAL"].dt.date.unique())


def run(paths, output_dir, master_path=None, tanggal=None, log=print):
    raws = [p for p in paths if p != master_path and p.lower().endswith((".csv", ".xlsx", ".xls"))]
    if not raws:
        raise FileNotFoundError("File Detail_Data_*.csv / .xlsx tidak ditemukan.")
    raw = next((p for p in raws if "detail" in os.path.basename(p).lower()), raws[0])
    log("File mentah :", os.path.basename(raw))
    log("Master MDS  :", os.path.basename(master_path) if master_path else f"bawaan ({len(DEFAULT_MASTER)} toko)")
    xlsx, res = generate(raw, master_path, tanggal)
    sm = res["summary"]
    log(f"Tanggal report: {res['tgl']:%d-%m-%Y} | {len(sm)} toko | "
        f"Total Net Sales MTD Rp {sm['M_SALES'].sum():,.0f}")
    if res["missing"]:
        log(f"[Peringatan] Toko di master tapi tidak ada di data: {', '.join(res['missing'])}")
    out = os.path.join(output_dir, output_name(res))
    with open(out, "wb") as f:
        f.write(xlsx)
    preview = sm.rename(columns={
        "D_JHK": "JHK", "D_SPD": "SPD", "D_SPD_TAGI": "SPD Tag I", "D_STD": "STD",
        "D_APC": "APC", "D_GM": "%GM", "M_JHK": "JHK MTD", "M_SALES": "Net Sales MTD",
        "M_SALES_TAGI": "Net Sales Tag I MTD", "M_SPD": "SPD MTD", "M_SPD_TAGI": "SPD Tag I MTD",
        "M_STD": "STD MTD", "M_APC": "APC MTD", "M_GM": "%GM MTD"})
    return out, preview, []
