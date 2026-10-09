"""Penyimpanan riwayat hasil olah data Darkstore (dipakai app olah data & dashboard).

Alur:
  app Streamlit (olah data) --save_all()--> penyimpanan --load_history()--> dashboard

Yang disimpan = ringkasan per toko per tanggal (bukan file mentah), dengan kunci
KD_STORE + TANGGAL. Upload ulang file yang sama TIDAK membuat data dobel: baris
dengan kunci yang sama ditimpa (upsert).

Backend:
  - SheetsBackend : Google Sheets (untuk Streamlit Cloud, data permanen)
  - LocalBackend  : folder CSV (untuk uji coba / jalan di komputer sendiri)
"""
import os

import numpy as np
import pandas as pd

HISTORY_SHEET = "history"
OOS_SHEET = "oos"

HISTORY_KEYS = ["KD_STORE", "TANGGAL"]
OOS_KEYS = ["KD_STORE", "PERIODE_END"]

HISTORY_COLS = ["KD_STORE", "NAMA_STORE", "NAMA_BRANCH", "TANGGAL",
                "JHK", "SALES", "SALES_TAGI", "SPD", "STD", "APC",
                "PERCENT_GM", "PCT_OOS_OFMB"]
OOS_COLS = ["KD_STORE", "PERIODE_END", "% OOS OFMB", "% OOS TAG I", "% OOS TAG K"]

# kolom teks; sisanya dianggap angka saat dibaca
STR_COLS = {"KD_STORE", "NAMA_STORE", "NAMA_BRANCH", "TANGGAL", "PERIODE_END"}


# ---------------------------------------------------------------- bentuk data
def build_history(df_detail):
    """Detail Data TRX -> tabel riwayat (1 baris = 1 toko x 1 tanggal)."""
    cols = [c for c in HISTORY_COLS if c in df_detail.columns]
    h = df_detail[cols].copy()
    h["TANGGAL"] = pd.to_datetime(h["TANGGAL"]).dt.strftime("%Y-%m-%d")
    h["KD_STORE"] = h["KD_STORE"].astype(str).str.strip()
    return h.drop_duplicates(HISTORY_KEYS, keep="last").reset_index(drop=True)


def build_oos(df_oos, periode_end):
    """OOS periode (Full/MTD) -> tabel riwayat OOS, diberi tanggal akhir periodenya."""
    o = df_oos.rename(columns={df_oos.columns[0]: "KD_STORE"}).copy()
    o["KD_STORE"] = o["KD_STORE"].astype(str).str.strip()
    o["PERIODE_END"] = pd.Timestamp(periode_end).strftime("%Y-%m-%d")
    cols = [c for c in OOS_COLS if c in o.columns]
    return o[cols].drop_duplicates(OOS_KEYS, keep="last").reset_index(drop=True)


def _clean(df):
    """Rapikan tipe data setelah dibaca dari penyimpanan."""
    if df.empty:
        return df
    for c in df.columns:
        if c in STR_COLS:
            df[c] = df[c].astype(str).str.strip()
        else:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


# ---------------------------------------------------------------- backend
class LocalBackend:
    """Simpan sebagai CSV di sebuah folder (uji coba / jalan lokal)."""

    def __init__(self, folder):
        self.folder = folder
        os.makedirs(folder, exist_ok=True)

    def _path(self, name):
        return os.path.join(self.folder, f"{name}.csv")

    def read(self, name):
        p = self._path(name)
        if not os.path.exists(p):
            return pd.DataFrame()
        return _clean(pd.read_csv(p, dtype=str).replace({"": np.nan}))

    def write(self, name, df):
        df.to_csv(self._path(name), index=False)


class SheetsBackend:
    """Simpan ke Google Sheets. Satu worksheet per tabel ('history', 'oos')."""

    CHUNK = 5000  # baris per request tulis

    def __init__(self, service_account_info, spreadsheet_id):
        import gspread  # import di sini supaya modul ini tetap bisa dipakai tanpa gspread
        self._gc = gspread.service_account_from_dict(dict(service_account_info))
        self._sh = self._gc.open_by_key(spreadsheet_id)

    def _ws(self, name, create=False):
        import gspread
        try:
            return self._sh.worksheet(name)
        except gspread.WorksheetNotFound:
            if not create:
                return None
            return self._sh.add_worksheet(title=name, rows=2, cols=20)

    def read(self, name):
        ws = self._ws(name)
        if ws is None:
            return pd.DataFrame()
        # UNFORMATTED_VALUE = angka asli (tanpa pembulatan tampilan Sheets)
        values = ws.get_all_values(value_render_option="UNFORMATTED_VALUE")
        if len(values) < 2:
            return pd.DataFrame()
        df = pd.DataFrame(values[1:], columns=values[0]).replace({"": np.nan})
        return _clean(df)

    def write(self, name, df):
        ws = self._ws(name, create=True)
        out = df.astype(object).where(df.notna(), "")
        rows = [list(map(str, out.columns))] + out.values.tolist()
        ws.clear()
        ws.resize(rows=max(len(rows), 2), cols=max(len(out.columns), 1))
        for i in range(0, len(rows), self.CHUNK):
            ws.update(rows[i:i + self.CHUNK], range_name=f"A{i + 1}", value_input_option="RAW")


def get_backend_from_streamlit():
    """Baca konfigurasi dari st.secrets. Return None kalau belum diatur.

    Isi secrets (lihat SETUP.md):
        [gcp_service_account]  ... (isi file JSON service account)
        [storage]
        spreadsheet_id = "..."
    """
    try:
        import streamlit as st
        sa = st.secrets["gcp_service_account"]
        sid = st.secrets["storage"]["spreadsheet_id"]
    except Exception:
        return None
    return SheetsBackend(sa, sid)


# ---------------------------------------------------------------- simpan / baca
def upsert(backend, name, new, keys):
    """Gabungkan data baru ke penyimpanan; baris dengan kunci sama ditimpa.
    Return (jumlah baris baru dikirim, total baris setelah digabung)."""
    old = backend.read(name)
    merged = new if old.empty else pd.concat([old, new], ignore_index=True)
    merged = merged.drop_duplicates(keys, keep="last")
    merged = merged.sort_values(keys).reset_index(drop=True)
    backend.write(name, merged)
    return len(new), len(merged)


def save_all(df_detail, df_oos, backend, log=print):
    """Simpan hasil olah data (riwayat per toko per tanggal + OOS periode)."""
    hist = build_history(df_detail)
    n, total = upsert(backend, HISTORY_SHEET, hist, HISTORY_KEYS)
    log(f"[Penyimpanan] history: {n} baris diproses, total tersimpan {total}")

    periode_end = pd.to_datetime(df_detail["TANGGAL"]).max()
    oos = build_oos(df_oos, periode_end)
    n, total = upsert(backend, OOS_SHEET, oos, OOS_KEYS)
    log(f"[Penyimpanan] oos: {n} baris diproses, total tersimpan {total}")


def save_if_configured(df_detail, df_oos, log=print, backend=None):
    """Dipanggil dari run(): simpan kalau penyimpanan sudah diatur.
    TIDAK PERNAH melempar error, supaya pembuatan report tidak ikut gagal."""
    try:
        backend = backend or get_backend_from_streamlit()
        if backend is None:
            log("[Penyimpanan] dilewati (secrets belum diatur)")
            return False
        save_all(df_detail, df_oos, backend, log)
        return True
    except Exception as e:  # noqa: BLE001
        log(f"[Penyimpanan] GAGAL, report tetap dibuat: {type(e).__name__}: {e}")
        return False


def load_history(backend):
    h = backend.read(HISTORY_SHEET)
    if h.empty:
        return h
    h["TANGGAL"] = pd.to_datetime(h["TANGGAL"])
    return h


def load_oos(backend):
    return backend.read(OOS_SHEET)


# ======================================================================
# Growth MDS (tambahan): snapshot report MDS per toko per tanggal
# ======================================================================
MDS_SHEET = "mds"
MDS_KEYS = ["KD_STORE", "TANGGAL"]


def build_mds(res):
    """Hasil mds_report.build() -> tabel snapshot (1 baris = 1 toko MDS x 1 tanggal)."""
    s = res["summary"].rename(columns={"Kode Toko": "KD_STORE",
                                       "Nama Toko": "NAMA_STORE",
                                       "Cabang": "NAMA_BRANCH"}).copy()
    s["KD_STORE"] = s["KD_STORE"].astype(str).str.strip()
    s["TANGGAL"] = pd.Timestamp(res["tgl"]).strftime("%Y-%m-%d")
    first = ["KD_STORE", "NAMA_STORE", "NAMA_BRANCH", "TANGGAL"]
    rest = [c for c in s.columns if c not in first]
    return s[first + rest].drop_duplicates(MDS_KEYS, keep="last").reset_index(drop=True)


def save_mds(res, backend, log=print):
    n, total = upsert(backend, MDS_SHEET, build_mds(res), MDS_KEYS)
    log(f"[Penyimpanan] mds: {n} baris diproses, total tersimpan {total}")
    n, total = upsert(backend, MDS_HIST_SHEET, build_mds_history(res), HISTORY_KEYS)
    log(f"[Penyimpanan] mds_history: {n} baris diproses, total tersimpan {total}")
    sla = build_mds_sla(res)
    if sla is not None:
        n, total = upsert(backend, MDS_SLA_SHEET, sla, MDS_SLA_KEYS)
        log(f"[Penyimpanan] mds_sla: {n} baris diproses, total tersimpan {total}")


def save_mds_if_configured(res, log=print, backend=None):
    """Simpan snapshot MDS kalau penyimpanan sudah diatur. TIDAK PERNAH melempar error."""
    try:
        backend = backend or get_backend_from_streamlit()
        if backend is None:
            log("[Penyimpanan] MDS dilewati (secrets belum diatur)")
            return False
        save_mds(res, backend, log)
        return True
    except Exception as e:  # noqa: BLE001
        log(f"[Penyimpanan] MDS GAGAL, report tetap dibuat: {type(e).__name__}: {e}")
        return False


def load_mds(backend):
    m = backend.read(MDS_SHEET)
    if m.empty:
        return m
    m["TANGGAL"] = pd.to_datetime(m["TANGGAL"])
    return m


# ---- riwayat harian MDS (semua tanggal di file TRX/SLA)
MDS_HIST_SHEET = "mds_history"
MDS_SLA_SHEET = "mds_sla"
MDS_SLA_KEYS = ["KD_STORE", "TANGGAL"]


def build_mds_history(res):
    """Data TRX toko MDS (semua tanggal) -> format sama dengan tab history."""
    return build_history(res["data_mds"])


def build_mds_sla(res):
    """Data SLA toko MDS (semua tanggal) -> ringkasan per toko per tanggal. None kalau tidak ada SLA."""
    s = res.get("sla_mds")
    if s is None or len(s) == 0:
        return None
    o = s.copy()
    o["KD_STORE"] = o["KD_STORE"].astype(str).str.strip()
    o["TANGGAL"] = pd.to_datetime(o["TANGGAL"]).dt.strftime("%Y-%m-%d")
    val = ["DELIVERY_ONTIME", "DELIVERY_LATE", "JUMLAH_DELIVERY"]
    keys = ["KD_STORE", "NAMA_STORE", "NAMA_BRANCH", "TANGGAL"]
    return o.groupby(keys, as_index=False)[val].sum().sort_values(MDS_SLA_KEYS).reset_index(drop=True)


def load_mds_history(backend):
    h = backend.read(MDS_HIST_SHEET)
    if h.empty:
        return h
    h["TANGGAL"] = pd.to_datetime(h["TANGGAL"])
    return h


def load_mds_sla(backend):
    s = backend.read(MDS_SLA_SHEET)
    if s.empty:
        return s
    s["TANGGAL"] = pd.to_datetime(s["TANGGAL"])
    return s


# ======================================================================
# OTD / LATE harian DS (tambahan): dari file SLA .csv (per toko per tanggal)
# ======================================================================
DS_SLA_DAILY_SHEET = "ds_sla_daily"
DS_SLA_DAILY_KEYS = ["KD_STORE", "TANGGAL"]


def build_ds_sla_daily(csv_path):
    """File Detail Data SLA (CSV) -> ringkasan per toko DS per tanggal."""
    s = pd.read_csv(csv_path, encoding="utf-8-sig", dtype={"KD_STORE": str})
    s.columns = [str(c).strip() for c in s.columns]
    if "NAMA_STORE" in s.columns:
        s = s[s["NAMA_STORE"].astype(str).str.upper().str.startswith("DS")]
    val = ["DELIVERY_ONTIME", "DELIVERY_LATE", "JUMLAH_DELIVERY"]
    for c in val:
        s[c] = pd.to_numeric(s[c], errors="coerce").fillna(0)
    s["KD_STORE"] = s["KD_STORE"].astype(str).str.strip()
    s["TANGGAL"] = pd.to_datetime(s["TANGGAL"], errors="coerce").dt.strftime("%Y-%m-%d")
    s = s.dropna(subset=["TANGGAL"])
    out = s.groupby(DS_SLA_DAILY_KEYS, as_index=False)[val].sum()
    return out.sort_values(DS_SLA_DAILY_KEYS).reset_index(drop=True)


def save_ds_sla_daily_if_configured(csv_path, log=print, backend=None):
    """Simpan OTD/LATE harian DS kalau penyimpanan sudah diatur. TIDAK PERNAH melempar error."""
    try:
        backend = backend or get_backend_from_streamlit()
        if backend is None:
            log("[Penyimpanan] DS SLA harian dilewati (secrets belum diatur)")
            return False
        n, total = upsert(backend, DS_SLA_DAILY_SHEET, build_ds_sla_daily(csv_path), DS_SLA_DAILY_KEYS)
        log(f"[Penyimpanan] ds_sla_daily: {n} baris diproses, total tersimpan {total}")
        return True
    except Exception as e:  # noqa: BLE001
        log(f"[Penyimpanan] DS SLA harian GAGAL, report tetap dibuat: {type(e).__name__}: {e}")
        return False


def load_ds_sla_daily(backend):
    s = backend.read(DS_SLA_DAILY_SHEET)
    if s.empty:
        return s
    s["TANGGAL"] = pd.to_datetime(s["TANGGAL"])
    return s
