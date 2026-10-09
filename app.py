"""Report Automation Darkstore Alfagift — web app (Streamlit).

Jalankan:  streamlit run app.py
"""
import datetime as dt
import os
import shutil
import tempfile
import traceback
from zoneinfo import ZoneInfo

import streamlit as st

from core import apo_darkstore, darkstore_ds, inventory_report, master_produk, mds_report, mtd_performance
from core.common import Log, find_soffice, prepare_inputs

st.set_page_config(page_title="Report Automation Darkstore", page_icon="📊", layout="wide")

WIB = ZoneInfo("Asia/Jakarta")
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


# ------------------------------------------------------------------ helpers
def save_uploads(uploaded_files, folder):
    os.makedirs(folder, exist_ok=True)
    for f in uploaded_files:
        with open(os.path.join(folder, os.path.basename(f.name)), "wb") as out:
            out.write(f.getbuffer())
    return prepare_inputs(folder)


def run_job(key, fn):
    """Jalankan fn(input_dir, output_dir, log) lalu simpan hasilnya di session_state."""
    log = Log()
    work = tempfile.mkdtemp(prefix="report_")
    in_dir, out_dir = os.path.join(work, "in"), os.path.join(work, "out")
    os.makedirs(out_dir)
    try:
        with st.spinner("Memproses..."):
            out, preview, extras = fn(in_dir, out_dir, log)
        files = []
        for p in [out] + list(extras):
            with open(p, "rb") as fh:
                files.append((os.path.basename(p), fh.read()))
        st.session_state[key] = {"ok": True, "files": files, "preview": preview, "log": log.text()}
    except Exception as e:  # noqa: BLE001
        st.session_state[key] = {"ok": False, "error": f"{type(e).__name__}: {e}",
                                 "trace": traceback.format_exc(), "log": log.text()}
    finally:
        shutil.rmtree(work, ignore_errors=True)


def show_result(key):
    res = st.session_state.get(key)
    if not res:
        return
    if res["ok"]:
        st.success("Selesai! Silakan download hasilnya.")
        cols = st.columns(len(res["files"]))
        for col, (name, data) in zip(cols, res["files"]):
            mime = "image/png" if name.endswith(".png") else XLSX_MIME
            col.download_button(f"⬇️ {name}", data, file_name=name, mime=mime,
                                use_container_width=True, key=f"dl_{key}_{name}")
        for name, data in res["files"]:
            if name.endswith(".png"):
                st.image(data, caption=name)
        if res.get("preview") is not None and len(res["preview"]):
            with st.expander("Preview data", expanded=False):
                st.dataframe(res["preview"], use_container_width=True, hide_index=True)
    else:
        st.error(res["error"])
        with st.expander("Detail error"):
            st.code(res["trace"])
    if res.get("log"):
        with st.expander("Log proses", expanded=not res["ok"]):
            st.code(res["log"], language=None)


def uploader(label, types, key, multiple=True, help=None):
    files = st.file_uploader(label, type=types, accept_multiple_files=multiple, key=key, help=help)
    if files is None:
        return []
    return files if multiple else [files]


# ------------------------------------------------------------------ halaman
def page_home():
    st.title("📊 Report Automation Darkstore Alfagift")
    st.write("Upload data mentah, klik **Proses**, lalu download laporan Excel yang sudah terformat. "
             "Pilih jenis laporan di sidebar.")
    rows = [(cat, title, desc) for cat, items in PAGES.items() if cat != "Beranda"
            for title, (_, desc) in items.items()]
    st.table({"Kategori": [r[0] for r in rows], "Laporan": [r[1] for r in rows],
              "Fungsi": [r[2] for r in rows]})
    if find_soffice() is None:
        st.info("LibreOffice tidak terdeteksi. Laporan tetap dibuat, hanya saja sel formula "
                "baru terisi nilainya saat file dibuka di Excel.")


def page_daily_ds():
    st.header("Report Daily Performance DS")
    st.caption("Input: `Detail_Data_....csv` + `[OOS] By Toko.xlsx` (boleh file satu-satu atau 1 ZIP).")
    files = uploader("Upload file / ZIP", ["csv", "xlsx", "zip"], "up_daily_ds")
    n_fb = st.number_input("Fallback hari untuk kolom Daily", 1, 7, darkstore_ds.N_FALLBACK_DAYS)
    if st.button("Proses", type="primary", disabled=not files, key="btn_daily_ds"):
        run_job("res_daily_ds", lambda i, o, log: darkstore_ds.run(
            save_uploads(files, i), o, with_delivery=False, n_fallback_days=int(n_fb), log=log))
    show_result("res_daily_ds")


def page_mtd_ds():
    st.header("MTD DS Performance + Delivery")
    st.caption("Sama seperti Daily DS, ditambah %Ontime & %Late dari Report Summary Dashboard.")
    c1, c2 = st.columns(2)
    with c1:
        files = uploader("Detail Data (.csv) + [OOS] By Toko (.xlsx) / ZIP", ["csv", "xlsx", "zip"], "up_mtd_ds")
    with c2:
        rs = uploader("Report Summary Dashboard (.xlsx)", ["xlsx"], "up_mtd_rs", multiple=False)
    n_fb = st.number_input("Fallback hari untuk kolom Daily", 1, 7, darkstore_ds.N_FALLBACK_DAYS, key="fb_mtd")

    def job(i, o, log):
        paths = save_uploads(files, i)
        rs_path = save_uploads(rs, os.path.join(i, "_rs"))[0]
        return darkstore_ds.run(paths, o, with_delivery=True, report_summary_path=rs_path,
                                n_fallback_days=int(n_fb), log=log)

    if st.button("Proses", type="primary", disabled=not (files and rs), key="btn_mtd_ds"):
        run_job("res_mtd_ds", job)
    show_result("res_mtd_ds")


def page_mds():
    st.header("Report Daily Mini Darkstore (MDS)")
    st.caption("Upload **1 ZIP** berisi `Detail_Data_*TRX.csv`, `Detail_Data_*SLA.csv` dan `Master MDS.xlsx` "
               "(file `Report Daily MDS ...xlsx` lama boleh ikut di dalam ZIP, otomatis diabaikan). "
               "Bisa juga upload file-file itu satu-satu. Hasil: `Report Daily MDS <tanggal>.xlsx`.")
    files = uploader("Upload ZIP / file", ["zip", "csv", "xlsx"], "up_mds", multiple=True)

    inp, tgl = None, None
    if files:
        sig = tuple((f.name, f.size) for f in files)
        cache = st.session_state.get("mds_inputs")
        if not cache or cache["sig"] != sig:
            try:
                tmp = tempfile.mkdtemp(prefix="mds_in_")
                paths = save_uploads(files, tmp)
                loaded = mds_report.load_inputs(paths)
                st.session_state["mds_inputs"] = {"sig": sig, "inp": loaded, "tmp": tmp}
            except Exception as e:  # noqa: BLE001
                st.session_state.pop("mds_inputs", None)
                st.error(f"File tidak bisa dibaca: {e}")
        cache = st.session_state.get("mds_inputs")
        if cache and cache["sig"] == sig:
            inp = cache["inp"]
            g = inp["files"]
            nm = lambda p: os.path.basename(p) if p else None  # noqa: E731
            c1, c2, c3 = st.columns(3)
            c1.markdown(f"**TRX**  \n{nm(g['trx']) or '❌ tidak ditemukan'}")
            c2.markdown(f"**SLA**  \n{nm(g['sla']) or '⚠️ tidak ada (%OTD & %LATE kosong)'}")
            c3.markdown(f"**Master MDS**  \n{nm(inp['master_path']) or 'bawaan (52 toko)'}")
            dates = mds_report.available_dates(inp)
            if dates:
                tgl = st.selectbox("Tanggal report", dates[::-1],
                                   format_func=lambda d: f"{d.day} {mds_report.BULAN[d.month]} {d.year}")
            else:
                st.error("Tidak ada tanggal yang sama antara file TRX dan SLA.")

    with st.expander("Master MDS terpisah (opsional)"):
        master = uploader("Master MDS.xlsx (menggantikan yang ada di ZIP)", ["xlsx", "csv"], "up_mds_master",
                          multiple=False)
    inc_raw = st.checkbox("Sertakan sheet Data Raw & Data SLA penuh (file jadi besar & lebih lama)", value=False)

    def job(i, o, log):
        paths = save_uploads(files, i)
        m = save_uploads(master, os.path.join(i, "_master"))[0] if master else None
        return mds_report.run(paths, o, master_path=m, tanggal=tgl, include_raw=inc_raw, log=log)

    ready = bool(files) and inp is not None and inp["trx"] is not None and tgl is not None
    if st.button("Proses", type="primary", disabled=not ready, key="btn_mds"):
        run_job("res_mds", job)
    show_result("res_mds")


def page_inventory():
    st.header("Report Inventory, OOS, MAT, Store Performance")
    st.caption("Upload 6 file mentah (satu-satu atau 1 ZIP). File Report periode sebelumnya opsional — "
               "kalau tidak ada, dipakai template bawaan.")
    with st.expander("Daftar file yang dibutuhkan"):
        for k, label in inventory_report.LABELS.items():
            st.markdown(f"- {label}")
    files = uploader("Upload file / ZIP", ["xlsx", "zip"], "up_inv")
    if st.button("Proses", type="primary", disabled=not files, key="btn_inv"):
        run_job("res_inv", lambda i, o, log: inventory_report.run(save_uploads(files, i), o, log=log))
    show_result("res_inv")


def page_mtd_perf():
    st.header("Report Performance Darkstore MTD")
    st.caption("Input: 1 file data mentah (xlsx/xls, termasuk export HTML/XML/CSV yang berekstensi .xls). "
               "Diurutkan dari % ONTIME terendah, heatmap pada kolom delivery.")
    files = uploader("Upload data mentah", ["xlsx", "xls", "csv", "html", "xml"], "up_perf", multiple=False)
    c1, c2 = st.columns(2)
    tgl = c1.text_input("Tanggal laporan (untuk judul)",
                        mtd_performance.default_tanggal_label(dt.datetime.now(WIB).date()))
    sheet = c2.text_input("Nama sheet (kosongkan = sheet pertama)", "")
    if st.button("Proses", type="primary", disabled=not files, key="btn_perf"):
        run_job("res_perf", lambda i, o, log: mtd_performance.run(
            save_uploads(files, i), o, tanggal_laporan=tgl, sheet_name=sheet.strip() or None, log=log))
    show_result("res_perf")


def page_master_produk():
    st.header("Merge Master Produk + Filter TAG")
    st.caption("Upload semua file Report Master Group PLU (.xls tab-separated) atau 1 ZIP.")
    files = uploader("Upload file / ZIP", ["xls", "txt", "tsv", "zip"], "up_mp")
    tags = st.text_input("TAG yang disimpan (pisahkan koma)", ", ".join(master_produk.DEFAULT_TAGS))
    if st.button("Proses", type="primary", disabled=not files, key="btn_mp"):
        tag_list = [t.strip() for t in tags.split(",") if t.strip()]
        run_job("res_mp", lambda i, o, log: master_produk.run(save_uploads(files, i), o, tags=tag_list, log=log))
    show_result("res_mp")


def page_apo():
    st.header("Update APO Darkstore")
    st.caption("Input: `GLI - NEW SAPA Performance_ALFAGIFT DARK STORE_Tabel pivot.csv`. "
               "Hasil: Excel + gambar PNG siap kirim.")
    files = uploader("Upload CSV", ["csv"], "up_apo", multiple=False)
    now = dt.datetime.now(WIB)
    c1, c2 = st.columns(2)
    tgl = c1.date_input("Tanggal", now.date(), format="DD/MM/YYYY")
    jam = c2.time_input("Jam", now.time().replace(second=0, microsecond=0))
    if st.button("Proses", type="primary", disabled=not files, key="btn_apo"):
        run_job("res_apo", lambda i, o, log: apo_darkstore.run(
            save_uploads(files, i), o, tanggal=tgl.strftime("%d/%m/%Y"), jam=jam.strftime("%H.%M"), log=log))
    show_result("res_apo")


PAGES = {
    "Beranda": {"Beranda": (page_home, "")},
    "Daily": {
        "Daily Performance DS": (page_daily_ds, "Report Daily Performance Darkstore dari Detail Data + OOS"),
        "Daily Mini Darkstore (MDS)": (page_mds, "Report Daily MDS: harian + MTD per toko mini darkstore"),
        "Inventory, OOS, MAT": (page_inventory, "Report Inventory, OOS, MAT & Store Performance"),
    },
    "MTD": {
        "MTD DS + Delivery": (page_mtd_ds, "Rekap MTD darkstore + %Ontime/%Late delivery"),
        "Performance Darkstore MTD": (page_mtd_perf, "Rekap MTD, urut % on-time terendah + heatmap"),
    },
    "Master Data": {
        "Master Produk (TAG)": (page_master_produk, "Gabung master produk & filter TAG D/G/S/W/6"),
        "Update APO Darkstore": (page_apo, "Pivot status order APO per toko (Excel + PNG)"),
    },
}

with st.sidebar:
    st.markdown("### 📊 Darkstore Reports")
    options = [(cat, title) for cat, items in PAGES.items() for title in items]
    choice = st.radio("Menu", options, format_func=lambda x: x[1] if x[0] == "Beranda" else f"{x[0]} · {x[1]}",
                      label_visibility="collapsed")
    st.divider()
    st.caption("Data hanya diproses sementara di server ini dan tidak disimpan.")

PAGES[choice[0]][choice[1]][0]()
