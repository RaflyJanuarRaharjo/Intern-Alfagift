"""Report Performance Darkstore MTD.

Sumber: notebooks/mtd/mtd_firstman.ipynb (cell generator, dipindah apa adanya).

generate_report_darkstore.py
=============================
Mengubah data mentah "Performance Darkstore" menjadi laporan Excel yang
sudah diformat rapi: judul, header berwarna navy, number format per kolom,
heatmap warna (color-scale) pada 4 kolom delivery, diurutkan dari
% DELIVERY_ONTIME terendah ke tertinggi, freeze panes, dan kolom bantu
disembunyikan — persis seperti format "Report Performance Darkstore MTD".

CARA PAKAI
----------
    python generate_report_darkstore.py INPUT.xlsx OUTPUT.xlsx --tanggal "10 SEP 2026"

Kalau --tanggal tidak diisi, otomatis pakai tanggal hari ini.
Kalau --sheet-input tidak diisi, otomatis pakai sheet aktif pertama.

STRUKTUR DATA MENTAH YANG DIBUTUHKAN
-------------------------------------
File input harus punya baris header (boleh di baris mana pun di 10 baris
pertama) dengan kolom-kolom berikut (nama boleh huruf besar/kecil bebas,
spasi/underscore bebas):

    KD_STORE, NAMA_STORE, KD_BRANCH, NAMA_BRANCH, REMARK, JHK,
    SALES, STRUK, GM, PERCENT_GM, SPD, STD, APC

Untuk 4 kolom delivery, skrip ini menerima SALAH SATU dari dua opsi:

    OPSI A (disarankan) - data mentah berupa JUMLAH SHIPMENT, bukan persen:
        DELIVERY_ONTIME, DELIVERY_LATE, DELIVERY_NO_SLA, JUMLAH_SHIPMENT_TUNDA
        -> hasil akhir tetap berupa FORMULA (=IFERROR(...)), jadi kalau nanti
           salah satu angka jumlah diubah manual di Excel, persennya otomatis
           ikut ter-update.

    OPSI B - data sudah berupa PERSEN (pecahan 0-1 atau 0-100):
        PERCENT_DELIVERY_ONTIME, PERCENT_DELIVERY_LATE,
        PERCENT_DELIVERY_NO_SLA, PERCENT_JUMLAH_SHIPMENT_TUNDA
        -> nilai ditulis apa adanya (angka biasa, bukan formula).

Rumus yang dipakai (mengikuti file referensi):
    % ONTIME  = ONTIME  / (ONTIME + LATE + NO_SLA)
    % LATE    = LATE    / (ONTIME + LATE + NO_SLA)
    % NO_SLA  = NO_SLA  / (ONTIME + LATE + NO_SLA)
    % TUNDA   = TUNDA   / (ONTIME + LATE + NO_SLA)
    (catatan: JUMLAH_SHIPMENT_TUNDA TIDAK ikut jadi pembagi, sesuai file asli)

DEPENDENSI
----------
    pip install pandas openpyxl   (biasanya sudah tersedia)
"""

import datetime
import os
import re
import sys

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------
# 1. KONFIGURASI TAMPILAN (samakan dengan file referensi)
# --------------------------------------------------------------------------

SHEET_NAME_OUTPUT = "MTD Performance"
TITLE_PREFIX = "REPORT PERFORMANCE DARKSTORE MTD"

HEADER_FILL = PatternFill("solid", fgColor="FF1F4E78")
HEADER_FONT = Font(name="Arial", size=10, bold=True, color="FFFFFFFF")
TITLE_FONT = Font(name="Arial", size=14, bold=True)
DATA_FONT = Font(name="Arial", size=10)

THIN_GRAY = Side(style="thin", color="FFBFBFBF")
CELL_BORDER = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY)

ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
ALIGN_CENTER = Alignment(horizontal="center", vertical="center")

RED = "FFF8696B"
YELLOW = "FFFFEB84"
GREEN = "FF63BE7B"

ROW_HEIGHT_TITLE = 21.75
ROW_HEIGHT_HEADER = 30
ROW_HEIGHT_DATA = 15

# Definisi kolom output: (huruf_kolom, header_text, lebar, alignment, number_format)
# G..M = metrik, N..Q = persen delivery (diisi lewat formula terpisah di bawah)
OUTPUT_COLUMNS = [
    ("A", "KD_STORE", 10, ALIGN_LEFT, "General"),
    ("B", "NAMA_STORE", 20, ALIGN_LEFT, "General"),
    ("C", "KD_BRANCH", 10, ALIGN_LEFT, "General"),
    ("D", "NAMA_BRANCH", 14, ALIGN_LEFT, "General"),
    ("E", "REMARK", 11, ALIGN_LEFT, "General"),
    ("F", "JHK", 6, ALIGN_CENTER, "0"),
    ("G", "SALES", 15, ALIGN_CENTER, "#,##0"),
    ("H", "STRUK", 9, ALIGN_CENTER, "#,##0"),
    ("I", "GM", 14, ALIGN_CENTER, "#,##0"),
    ("J", "PERCENT_GM", 11, ALIGN_CENTER, "0.00%"),
    ("K", "SPD", 14, ALIGN_CENTER, "#,##0"),
    ("L", "STD", 9, ALIGN_CENTER, "0.00"),
    ("M", "APC", 13, ALIGN_CENTER, "#,##0.00"),
    ("N", "% DELIVERY_ONTIME", 15, ALIGN_CENTER, "0%"),
    ("O", "% DELIVERY_LATE", 13, ALIGN_CENTER, "0%"),
    ("P", "% DELIVERY_NO_SLA", 15, ALIGN_CENTER, "0%"),
    ("Q", "% JUMLAH_SHIPMENT_TUNDA", 20, ALIGN_CENTER, "0%"),
]
LAST_VISIBLE_COL = "Q"
# Kolom bantu (raw count) yang disembunyikan, mengikuti file referensi
HELPER_COLS = ["R", "S", "T", "U", "V", "W"]  # R,S sengaja dikosongkan (buffer)
COL_ONTIME, COL_LATE, COL_NOSLA, COL_TUNDA = "T", "U", "V", "W"

# --------------------------------------------------------------------------
# 2. PENCARIAN KOLOM DI FILE MENTAH (fleksibel: tidak peduli spasi/huruf besar-kecil)
# --------------------------------------------------------------------------

def _norm(s):
    return re.sub(r"[^A-Z0-9]", "", str(s).upper())


BASE_FIELD_ALIASES = {
    "KD_STORE": ["KDSTORE", "STORECODE", "KODESTORE"],
    "NAMA_STORE": ["NAMASTORE", "STORENAME"],
    "KD_BRANCH": ["KDBRANCH", "BRANCHCODE", "KODEBRANCH"],
    "NAMA_BRANCH": ["NAMABRANCH", "BRANCHNAME"],
    "REMARK": ["REMARK"],
    "JHK": ["JHK"],
    "SALES": ["SALES"],
    "STRUK": ["STRUK"],
    "GM": ["GM"],
    "PERCENT_GM": ["PERCENTGM", "PCTGM", "GMPERCENT"],
    "SPD": ["SPD"],
    "STD": ["STD"],
    "APC": ["APC"],
}

COUNT_FIELD_ALIASES = {
    "ONTIME": ["DELIVERYONTIME", "ONTIME", "JMLONTIME", "JUMLAHONTIME"],
    "LATE": ["DELIVERYLATE", "LATE", "JMLLATE", "JUMLAHLATE"],
    "NOSLA": ["DELIVERYNOSLA", "NOSLA", "JMLNOSLA", "JUMLAHNOSLA"],
    "TUNDA": ["JUMLAHSHIPMENTTUNDA", "SHIPMENTTUNDA", "TUNDA", "JMLTUNDA"],
}

PERCENT_FIELD_ALIASES = {
    "ONTIME": ["PERCENTDELIVERYONTIME", "PCTDELIVERYONTIME", "DELIVERYONTIMEPERCENT"],
    "LATE": ["PERCENTDELIVERYLATE", "PCTDELIVERYLATE", "DELIVERYLATEPERCENT"],
    "NOSLA": ["PERCENTDELIVERYNOSLA", "PCTDELIVERYNOSLA", "DELIVERYNOSLAPERCENT"],
    "TUNDA": ["PERCENTJUMLAHSHIPMENTTUNDA", "PCTSHIPMENTTUNDA", "SHIPMENTTUNDAPERCENT"],
}


def find_column(df, aliases):
    """Cari nama kolom asli di df yang cocok dengan salah satu alias (case/format bebas)."""
    norm_map = {_norm(c): c for c in df.columns}
    for alias in aliases:
        if alias in norm_map:
            return norm_map[alias]
    return None


def _to_float(v):
    """Konversi nilai (angka, string dengan koma ribuan, dsb) jadi float. Aman dari NaN/None/teks kosong."""
    if v is None:
        return 0.0
    if isinstance(v, float) and pd.isna(v):
        return 0.0
    if isinstance(v, str):
        v = v.strip().replace(",", "").rstrip("%").strip()
        if v == "" or v.upper() in ("NAN", "NONE", "-"):
            return 0.0
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _to_percent_fraction(v):
    """Konversi nilai persen (bisa '15.57%', 15.57, atau 0.1557) jadi pecahan 0-1."""
    had_percent_sign = isinstance(v, str) and "%" in v
    f = _to_float(v)
    if had_percent_sign or f > 1:
        return f / 100
    return f


def _try_read_spreadsheetml_xml(input_path):
    """
    Baca format 'SpreadsheetML' -- XML Excel 2003 lama (<?mso-application progid="Excel.Sheet"?>,
    root <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet">). Ini XML, BUKAN html,
    jadi pd.read_html tidak bisa mendeteksinya. Cukup umum jadi format export sistem lama/ASP.NET.
    """
    import xml.etree.ElementTree as ET

    ns = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}
    tree = ET.parse(input_path)
    root = tree.getroot()
    worksheet = root.find("ss:Worksheet", ns)
    if worksheet is None:
        raise ValueError("Bukan format SpreadsheetML (tidak ada elemen <Worksheet>).")
    table = worksheet.find("ss:Table", ns)
    if table is None:
        raise ValueError("Bukan format SpreadsheetML (tidak ada elemen <Table>).")

    rows_data = []
    for row in table.findall("ss:Row", ns):
        row_vals = []
        col_idx = 0
        for cell in row.findall("ss:Cell", ns):
            idx_attr = cell.get("{urn:schemas-microsoft-com:office:spreadsheet}Index")
            if idx_attr:
                target_idx = int(idx_attr) - 1
                while col_idx < target_idx:
                    row_vals.append(None)
                    col_idx += 1
            data_elem = cell.find("ss:Data", ns)
            row_vals.append(data_elem.text if data_elem is not None else None)
            col_idx += 1
        rows_data.append(row_vals)

    if not rows_data:
        raise ValueError("Tabel SpreadsheetML kosong.")
    maxlen = max(len(r) for r in rows_data)
    rows_data = [r + [None] * (maxlen - len(r)) for r in rows_data]
    return pd.DataFrame(rows_data)


def _try_read_delimited_text(input_path):
    """Coba baca sebagai teks berpemisah (CSV/TSV/semicolon/pipe)."""
    for sep in [None, ",", ";", "\t", "|"]:
        try:
            df = pd.read_csv(
                input_path, sep=sep, engine="python", header=None,
                encoding="utf-8-sig", on_bad_lines="skip",
            )
            if df.shape[1] > 1:
                return df
        except Exception:
            continue
    raise ValueError("Tidak bisa dibaca sebagai file teks berpemisah (CSV/TSV/dll).")


def _read_raw_table(input_path, sheet_name=None):
    """
    Baca file input sebagai tabel mentah TANPA asumsi baris header (header=None).
    Otomatis mencoba beberapa format secara berurutan, karena file "Excel" hasil
    export dashboard/BI internal sering kali BUKAN file Excel biner asli:

        1. Excel biner asli (.xlsx via openpyxl / .xls via xlrd)
        2. Tabel HTML yang diberi ekstensi .xls/.xlsx (paling umum)
        3. SpreadsheetML -- XML Excel 2003 lama
        4. Teks berpemisah (CSV/TSV/semicolon/pipe) yang diberi ekstensi .xls/.xlsx
    """
    attempts = [
        ("Excel biner (.xlsx/.xls)",
         lambda: pd.read_excel(input_path, sheet_name=sheet_name if sheet_name else 0, header=None)),
        ("Tabel HTML", lambda: _pick_best_html_table(input_path)),
        ("SpreadsheetML XML", lambda: _try_read_spreadsheetml_xml(input_path)),
        ("Teks berpemisah (CSV/TSV)", lambda: _try_read_delimited_text(input_path)),
    ]

    errors = []
    for label, fn in attempts:
        try:
            df = fn()
            if label != "Excel biner (.xlsx/.xls)":
                print(f"Catatan: file terbaca sebagai format '{label}' (bukan file Excel biner asli).")
            return df
        except Exception as e:
            errors.append(f"- {label}: {e}")

    raise ValueError(
        "File tidak bisa dibaca dengan format apa pun yang didukung. Sudah dicoba:\n"
        + "\n".join(errors)
        + "\n\nCoba buka file ini secara manual (Excel/Google Sheets), lalu 'Save As'/'Export' "
        "ulang sebagai file Excel (.xlsx) yang baru, kemudian upload ulang."
    )


def _pick_best_html_table(input_path):
    tables = pd.read_html(input_path, header=None)
    # Pilih tabel yang mengandung 'KD_STORE' di salah satu selnya; kalau tidak ketemu,
    # pilih tabel dengan jumlah sel terbanyak (kemungkinan besar itu tabel utama).
    for t in tables:
        flat = [_norm(v) for row in t.values.tolist() for v in row]
        if "KDSTORE" in flat:
            return t
    return max(tables, key=lambda t: t.shape[0] * t.shape[1])


def load_raw_data(input_path, sheet_name=None):
    """Baca file mentah & cari otomatis baris header (yang mengandung KD_STORE) di 10 baris awal."""
    raw = _read_raw_table(input_path, sheet_name)

    # Kasus file HTML (hasil fallback dari _read_raw_table): pandas.read_html otomatis
    # mendeteksi baris <th> sebagai nama kolom, jadi header-nya sudah benar dari sananya.
    if find_column(raw, BASE_FIELD_ALIASES["KD_STORE"]) is not None:
        return raw.reset_index(drop=True)

    header_row_idx = None
    for i in range(min(10, len(raw))):
        row_vals = [_norm(v) for v in raw.iloc[i].tolist()]
        if "KDSTORE" in row_vals:
            header_row_idx = i
            break
    if header_row_idx is None:
        raise ValueError(
            "Tidak menemukan kolom KD_STORE di 10 baris pertama file input. "
            "Pastikan file mentah punya header dengan kolom KD_STORE."
        )

    header = raw.iloc[header_row_idx].tolist()
    df = raw.iloc[header_row_idx + 1:].copy()
    df.columns = header
    df = df.reset_index(drop=True)
    # Buang baris yang sepenuhnya kosong (kadang ada baris kosong sisa dari tabel HTML)
    df = df.dropna(how="all").reset_index(drop=True)
    return df


# --------------------------------------------------------------------------
# 3. PROSES DATA
# --------------------------------------------------------------------------

def build_dataset(df):
    """Kembalikan (list of dict data dasar, mode_delivery, kolom_delivery_terpakai)."""
    resolved = {}
    missing = []
    for field, aliases in BASE_FIELD_ALIASES.items():
        col = find_column(df, aliases)
        if col is None:
            missing.append(field)
        resolved[field] = col
    if missing:
        raise ValueError(
            "Kolom berikut tidak ditemukan di file input: "
            + ", ".join(missing)
            + ". Cek kembali nama header di file mentah."
        )

    # Cari kolom delivery: coba mode "jumlah" dulu, kalau tidak lengkap coba mode "persen"
    count_cols = {k: find_column(df, v) for k, v in COUNT_FIELD_ALIASES.items()}
    percent_cols = {k: find_column(df, v) for k, v in PERCENT_FIELD_ALIASES.items()}

    if all(count_cols.values()):
        mode = "count"
        delivery_cols = count_cols
    elif all(percent_cols.values()):
        mode = "percent"
        delivery_cols = percent_cols
    else:
        raise ValueError(
            "Kolom data delivery tidak lengkap. Sediakan salah satu dari dua opsi:\n"
            "  OPSI A (jumlah): DELIVERY_ONTIME, DELIVERY_LATE, DELIVERY_NO_SLA, "
            "JUMLAH_SHIPMENT_TUNDA\n"
            "  OPSI B (persen): PERCENT_DELIVERY_ONTIME, PERCENT_DELIVERY_LATE, "
            "PERCENT_DELIVERY_NO_SLA, PERCENT_JUMLAH_SHIPMENT_TUNDA"
        )

    NUMERIC_FIELDS = {"JHK", "SALES", "STRUK", "GM", "SPD", "STD", "APC"}

    rows = []
    for _, r in df.iterrows():
        item = {}
        for field, col in resolved.items():
            val = r[col]
            if field in NUMERIC_FIELDS:
                item[field] = _to_float(val)
            elif field == "PERCENT_GM":
                item[field] = _to_percent_fraction(val)
            else:
                item[field] = val

        if mode == "count":
            ontime = _to_float(r[delivery_cols["ONTIME"]])
            late = _to_float(r[delivery_cols["LATE"]])
            nosla = _to_float(r[delivery_cols["NOSLA"]])
            tunda = _to_float(r[delivery_cols["TUNDA"]])
            item["_ontime_count"] = ontime
            item["_late_count"] = late
            item["_nosla_count"] = nosla
            item["_tunda_count"] = tunda
            denom = ontime + late + nosla
            item["_sort_pct"] = (ontime / denom) if denom else 0
        else:
            item["_pct_ontime"] = _to_percent_fraction(r[delivery_cols["ONTIME"]])
            item["_pct_late"] = _to_percent_fraction(r[delivery_cols["LATE"]])
            item["_pct_nosla"] = _to_percent_fraction(r[delivery_cols["NOSLA"]])
            item["_pct_tunda"] = _to_percent_fraction(r[delivery_cols["TUNDA"]])
            item["_sort_pct"] = item["_pct_ontime"]
        rows.append(item)

    # Urutkan dari % DELIVERY_ONTIME terendah -> tertinggi (store paling bermasalah di atas)
    rows.sort(key=lambda x: x["_sort_pct"])
    return rows, mode


# --------------------------------------------------------------------------
# 4. TULIS WORKBOOK OUTPUT
# --------------------------------------------------------------------------

def write_report(rows, mode, output_path, report_date_label):
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME_OUTPUT

    n_rows = len(rows)
    first_data_row = 4
    last_data_row = first_data_row + n_rows - 1

    # --- Judul (baris 1) ---
    title_text = f"{TITLE_PREFIX} {report_date_label}"
    ws["A1"] = title_text
    ws["A1"].font = TITLE_FONT
    ws["A1"].alignment = ALIGN_LEFT
    ws.merge_cells(f"A1:{LAST_VISIBLE_COL}1")
    ws.row_dimensions[1].height = ROW_HEIGHT_TITLE

    # --- Header (baris 3) ---
    for col_letter, header_text, width, _align, _fmt in OUTPUT_COLUMNS:
        cell = ws[f"{col_letter}3"]
        cell.value = header_text
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.border = CELL_BORDER
        cell.alignment = ALIGN_CENTER
        ws.column_dimensions[col_letter].width = width
    ws.row_dimensions[3].height = ROW_HEIGHT_HEADER

    # --- Data (mulai baris 4) ---
    base_fields_order = [
        "KD_STORE", "NAMA_STORE", "KD_BRANCH", "NAMA_BRANCH", "REMARK",
        "JHK", "SALES", "STRUK", "GM", "PERCENT_GM", "SPD", "STD", "APC",
    ]
    for i, item in enumerate(rows):
        r = first_data_row + i
        for col_idx, field in enumerate(base_fields_order):
            col_letter = OUTPUT_COLUMNS[col_idx][0]
            cell = ws[f"{col_letter}{r}"]
            cell.value = item[field]
            cell.font = DATA_FONT
            cell.border = CELL_BORDER
            cell.alignment = OUTPUT_COLUMNS[col_idx][3]
            cell.number_format = OUTPUT_COLUMNS[col_idx][4]

        if mode == "count":
            ws[f"{COL_ONTIME}{r}"] = item["_ontime_count"]
            ws[f"{COL_LATE}{r}"] = item["_late_count"]
            ws[f"{COL_NOSLA}{r}"] = item["_nosla_count"]
            ws[f"{COL_TUNDA}{r}"] = item["_tunda_count"]
            ws[f"N{r}"] = (
                f"=IFERROR({COL_ONTIME}{r}/({COL_ONTIME}{r}+{COL_LATE}{r}+{COL_NOSLA}{r}),0)"
            )
            ws[f"O{r}"] = (
                f"=IFERROR({COL_LATE}{r}/({COL_ONTIME}{r}+{COL_LATE}{r}+{COL_NOSLA}{r}),0)"
            )
            ws[f"P{r}"] = (
                f"=IFERROR({COL_NOSLA}{r}/({COL_ONTIME}{r}+{COL_LATE}{r}+{COL_NOSLA}{r}),0)"
            )
            ws[f"Q{r}"] = (
                f"=IFERROR({COL_TUNDA}{r}/({COL_ONTIME}{r}+{COL_LATE}{r}+{COL_NOSLA}{r}),0)"
            )
        else:
            ws[f"N{r}"] = item["_pct_ontime"]
            ws[f"O{r}"] = item["_pct_late"]
            ws[f"P{r}"] = item["_pct_nosla"]
            ws[f"Q{r}"] = item["_pct_tunda"]

        for col_letter in ["N", "O", "P", "Q"]:
            cell = ws[f"{col_letter}{r}"]
            cell.font = DATA_FONT
            cell.border = CELL_BORDER
            cell.alignment = ALIGN_CENTER
            cell.number_format = "0%"

        ws.row_dimensions[r].height = ROW_HEIGHT_DATA

    # --- Sembunyikan kolom bantu ---
    for col_letter in HELPER_COLS:
        ws.column_dimensions[col_letter].hidden = True
        if col_letter not in ("R", "S"):
            ws.column_dimensions[col_letter].width = 13

    # --- Freeze panes: baris 1-3 & kolom A tetap terlihat saat scroll ---
    ws.freeze_panes = "B4"

    # --- Heatmap warna (color scale) untuk 4 kolom delivery ---
    # % ONTIME: makin tinggi makin baik -> merah (rendah) - kuning - hijau (tinggi)
    ws.conditional_formatting.add(
        f"N{first_data_row}:N{last_data_row}",
        ColorScaleRule(
            start_type="min", start_color=RED,
            mid_type="percentile", mid_value=50, mid_color=YELLOW,
            end_type="max", end_color=GREEN,
        ),
    )
    # % LATE, % NO_SLA, % TUNDA: makin rendah makin baik -> hijau (rendah) - kuning - merah (tinggi)
    for col_letter in ["O", "P", "Q"]:
        ws.conditional_formatting.add(
            f"{col_letter}{first_data_row}:{col_letter}{last_data_row}",
            ColorScaleRule(
                start_type="min", start_color=GREEN,
                mid_type="percentile", mid_value=50, mid_color=YELLOW,
                end_type="max", end_color=RED,
            ),
        )

    wb.save(output_path)


# --------------------------------------------------------------------------
# 5. ENTRY POINT (web)
# --------------------------------------------------------------------------

def default_tanggal_label(today=None):
    today = today or datetime.date.today()
    bulan = ["JAN", "FEB", "MAR", "APR", "MEI", "JUN", "JUL", "AGU", "SEP", "OKT", "NOV", "DES"]
    return f"{today.day} {bulan[today.month - 1]} {today.year}"


def run(paths, output_dir, tanggal_laporan=None, sheet_name=None, log=print):
    if not paths:
        raise FileNotFoundError("Upload 1 file data mentah.")
    input_path = paths[0]
    tanggal_laporan = (tanggal_laporan or default_tanggal_label()).strip()
    df = load_raw_data(input_path, sheet_name or None)
    log(f"{len(df)} baris data mentah terbaca dari {os.path.basename(input_path)}.")
    rows, mode = build_dataset(df)
    log("Mode data delivery: " + ("JUMLAH shipment (formula otomatis)" if mode == "count" else "PERSEN"))
    tgl_hari = tanggal_laporan.split()[0] if tanggal_laporan else ""
    name = f"Report_Performance_Darkstore_MTD {tgl_hari}.xlsx".replace("/", "-")
    out = os.path.join(output_dir, name)
    write_report(rows, mode, out, tanggal_laporan)
    log(f"Laporan selesai dibuat: {name} ({len(rows)} toko, diurutkan dari % ONTIME terendah)")
    preview = pd.DataFrame(rows)
    if not preview.empty:
        preview["% ONTIME"] = (preview["_sort_pct"] * 100).round(1)
        preview = preview[[c for c in preview.columns if not c.startswith("_")]]
    return out, preview, []
