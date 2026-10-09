"""Dashboard Performance Darkstore (read-only), dengan navbar: Growth DS | MDS | Compare Toko.

Deploy sebagai app Streamlit KEDUA:  Main file path = dashboard/app.py
Membaca riwayat yang disimpan oleh app olah data. Tidak ada upload di sini.
"""
import os
import sys

import pandas as pd
import streamlit as st

# supaya `import core...` jalan walau script ada di folder dashboard/
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core import storage  # noqa: E402
from core.growth import (compute_growth, has_previous_month, incomplete_dates,  # noqa: E402
                         last_complete_date)

st.set_page_config(page_title="Dashboard Darkstore", layout="wide")


# ---------------------------------------------------------------- akses
def check_password():
    """Gerbang password sederhana. Isi `dashboard_password` di Streamlit Secrets."""
    pw = st.secrets.get("dashboard_password", None)
    if not pw:
        st.error("Password dashboard belum diatur di Secrets (`dashboard_password`).")
        st.stop()
    if st.session_state.get("auth_ok"):
        return
    entered = st.text_input("Password", type="password")
    if entered and entered == pw:
        st.session_state["auth_ok"] = True
        st.rerun()
    elif entered:
        st.error("Password salah.")
    st.stop()


check_password()


# ---------------------------------------------------------------- data
@st.cache_data(ttl=300, show_spinner="Memuat data...")
def load():
    backend = storage.get_backend_from_streamlit()
    if backend is None:
        return None, None, None, None, None
    return (storage.load_history(backend), storage.load_oos(backend),
            storage.load_mds(backend), storage.load_mds_history(backend),
            storage.load_mds_sla(backend))


hist, oos, mds, mds_hist, mds_sla = load()
if hist is None:
    st.error("Penyimpanan belum diatur (secrets `gcp_service_account` dan `storage.spreadsheet_id`).")
    st.stop()


# ---------------------------------------------------------------- format angka (gaya Indonesia)
def fmt_num(v, d=0):
    """1234567.8 -> '1.234.568' (d=0) atau '1.234.567,8' (d=1)."""
    if v is None or pd.isna(v):
        return "-"
    s = f"{v:,.{d}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_pct(v, d=1):
    if v is None or pd.isna(v):
        return "-"
    return fmt_num(v, d) + "%"


def styled(df, int_cols=(), pct_cols=(), pct_dec=1):
    """Styler: tampilan pakai titik ribuan, nilai asli tetap angka (sort tetap benar)."""
    fm = {}
    for c in int_cols:
        if c in df.columns:
            fm[c] = lambda v: fmt_num(v, 0)
    for c in pct_cols:
        if c in df.columns:
            fm[c] = lambda v, d=pct_dec: fmt_pct(v, d)
    return df.style.format(fm, na_rep="-")


def reload_button(key):
    if st.button("Muat ulang data", key=key):
        st.cache_data.clear()
        st.rerun()


def pick_dates(bad_dates, dates, key):
    use_all = st.checkbox("Sertakan tanggal yang datanya belum lengkap", value=False, key=key)
    pick = dates if (use_all or not bad_dates) else [d for d in dates if pd.Timestamp(d) not in set(bad_dates)]
    return (pick or dates), use_all


def search_box(key):
    return st.text_input("Cari toko / kode / cabang", key=key,
                         placeholder="mis. 1A08, puskopkar, jambi").strip()


def apply_search(df, q):
    """Filter baris yang kode/nama toko/cabangnya mengandung teks pencarian.
    Pisahkan dengan koma untuk mencari beberapa sekaligus (mis. '1A08, 1A0Y')."""
    if not q or df is None or df.empty:
        return df
    cols = [c for c in ("KD_STORE", "NAMA_STORE", "NAMA_BRANCH") if c in df.columns]
    hay = df[cols].astype(str).agg(" ".join, axis=1).str.lower()
    mask = pd.Series(False, index=df.index)
    for term in [t.strip().lower() for t in q.split(",") if t.strip()]:
        mask |= hay.str.contains(term, regex=False)
    return df[mask]


GROWTH_COLS = ["KD_STORE", "NAMA_STORE", "NAMA_BRANCH", "Sales", "Growth DoD", "Growth WoW",
               "MTD", "MTD bulan lalu", "Growth MTD"]
GROWTH_INT = ["Sales", "MTD", "MTD bulan lalu"]
GROWTH_PCT = ["Growth DoD", "Growth WoW", "Growth MTD"]


# ================================================================ HALAMAN: GROWTH DS
def page_ds():
    st.title("Dashboard Performance Darkstore")
    if hist.empty:
        st.info("Belum ada data tersimpan. Jalankan app olah data dulu.")
        return

    with st.sidebar:
        st.header("Filter")
        reload_button("reload_ds")
        cabang = st.multiselect("Cabang", sorted(hist["NAMA_BRANCH"].dropna().unique()), key="cab_ds")
        q = search_box("q_ds")
        dates = sorted(hist["TANGGAL"].dt.normalize().unique())
        bad_dates = incomplete_dates(hist)
        pick, use_all = pick_dates(bad_dates, dates, "all_ds")
        as_of = st.selectbox("Per tanggal", pick[::-1], key="asof_ds",
                             format_func=lambda d: pd.Timestamp(d).strftime("%d %b %Y"))

    view = hist if not cabang else hist[hist["NAMA_BRANCH"].isin(cabang)]
    view = apply_search(view, q)
    if view.empty:
        st.info("Tidak ada toko yang cocok dengan pencarian.")
        return
    as_of = pd.Timestamp(as_of)
    g = compute_growth(view, as_of)

    if bad_dates and not use_all:
        st.warning("Tanggal berikut disembunyikan karena datanya tampak belum lengkap "
                   "(total Sales jauh di bawah hari-hari sebelumnya): "
                   + ", ".join(pd.Timestamp(d).strftime("%d %b") for d in bad_dates)
                   + ". Centang 'Sertakan tanggal...' di sidebar untuk menampilkannya.")

    month_start = as_of.replace(day=1)
    mtd = view[(view["TANGGAL"].dt.normalize() >= month_start) & (view["TANGGAL"].dt.normalize() <= as_of)]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric(f"Sales {as_of.strftime('%d %b')}", fmt_num(g["Sales"].sum()))
    c2.metric("Sales MTD", fmt_num(mtd["SALES"].sum()))
    c3.metric("Rata-rata SPD (MTD)", fmt_num(mtd["SPD"].mean()))
    c4.metric("Rata-rata %GM (MTD)", fmt_pct(mtd["PERCENT_GM"].mean() * 100, 2))

    st.subheader("Tren Sales harian")
    trend_src = view if use_all else view[~view["TANGGAL"].dt.normalize().isin(bad_dates)]
    trend = trend_src.groupby(trend_src["TANGGAL"].dt.normalize())["SALES"].sum()
    st.line_chart(trend)

    st.subheader("Performa per toko")
    if not has_previous_month(hist, as_of):
        st.caption("Growth MTD vs bulan lalu kosong: riwayat bulan sebelumnya belum tersimpan.")
    show = g.copy()
    for col in GROWTH_PCT:
        show[col] = show[col] * 100
    st.dataframe(styled(show[GROWTH_COLS], GROWTH_INT, GROWTH_PCT),
                 use_container_width=True, hide_index=True)

    left, right = st.columns(2)
    with left:
        st.subheader("5 toko Sales MTD tertinggi")
        st.bar_chart(g.nlargest(5, "MTD").set_index("NAMA_STORE")["MTD"])
    with right:
        st.subheader("5 toko Sales MTD terendah")
        st.bar_chart(g.nsmallest(5, "MTD").set_index("NAMA_STORE")["MTD"])

    if oos is not None and not oos.empty:
        st.subheader("OOS periode terbaru")
        latest = oos["PERIODE_END"].max()
        o = oos[oos["PERIODE_END"] == latest].merge(
            hist.drop_duplicates("KD_STORE", keep="last")[["KD_STORE", "NAMA_STORE", "NAMA_BRANCH"]],
            on="KD_STORE", how="left")
        if cabang:
            o = o[o["NAMA_BRANCH"].isin(cabang)]
        o = apply_search(o, q)
        st.caption(f"Periode s.d. {latest}")
        oos_cols = ["% OOS OFMB", "% OOS TAG I", "% OOS TAG K"]
        for col in oos_cols:
            if col in o.columns:
                o[col] = o[col] * 100
        st.dataframe(styled(o.sort_values("% OOS OFMB", ascending=False), pct_cols=oos_cols, pct_dec=2),
                     use_container_width=True, hide_index=True)


# ================================================================ HALAMAN: MDS
def page_mds():
    st.title("Dashboard MDS (Mini Darkstore)")

    if mds_hist is None or mds_hist.empty:
        st.info("Riwayat harian MDS belum tersimpan. Proses report MDS sekali lagi di app olah data "
                "(menu Daily Mini Darkstore) supaya semua tanggal di file tersimpan.")
        return

    with st.sidebar:
        st.header("Filter")
        reload_button("reload_mds")
        cabang = st.multiselect("Cabang", sorted(mds_hist["NAMA_BRANCH"].dropna().unique()), key="cab_mds")
        q = search_box("q_mds")
    mh = mds_hist if not cabang else mds_hist[mds_hist["NAMA_BRANCH"].isin(cabang)]
    mh = apply_search(mh, q)
    if mh.empty:
        st.info("Tidak ada toko MDS yang cocok dengan pencarian.")
        return
    with st.sidebar:
        dates = sorted(mh["TANGGAL"].dt.normalize().unique())
        bad_dates = incomplete_dates(mds_hist)
        pick, use_all = pick_dates(bad_dates, dates, "all_mds")
        as_of = st.selectbox("Per tanggal", pick[::-1], key="asof_mds",
                             format_func=lambda d: pd.Timestamp(d).strftime("%d %b %Y"))
    as_of = pd.Timestamp(as_of)

    if bad_dates and not use_all:
        st.warning("Tanggal berikut disembunyikan karena datanya tampak belum lengkap: "
                   + ", ".join(pd.Timestamp(d).strftime("%d %b") for d in bad_dates)
                   + ". Centang 'Sertakan tanggal...' di sidebar untuk menampilkannya.")

    # ---- KPI + growth
    gm = compute_growth(mh, as_of)
    prev = gm["MTD bulan lalu"].sum()
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Jumlah toko MDS", fmt_num(len(gm)))
    k2.metric(f"Sales MDS {as_of.strftime('%d %b')}", fmt_num(gm["Sales"].sum()))
    k3.metric("Sales MDS MTD", fmt_num(gm["MTD"].sum()))
    k4.metric("Growth MTD MDS", fmt_pct((gm["MTD"].sum() / prev - 1) * 100) if prev else "-")
    if not has_previous_month(mds_hist, as_of):
        st.caption("Growth MTD vs bulan lalu kosong: riwayat MDS bulan sebelumnya belum tersimpan.")

    st.subheader("Performa MDS per toko")
    show = gm.copy()
    for col in GROWTH_PCT:
        show[col] = show[col] * 100
    st.dataframe(styled(show[GROWTH_COLS], GROWTH_INT, GROWTH_PCT),
                 use_container_width=True, hide_index=True)

    # ---- tren harian
    st.subheader("Tren harian MDS")
    ts = mh if use_all else mh[~mh["TANGGAL"].dt.normalize().isin(bad_dates)]
    day = ts["TANGGAL"].dt.normalize()
    if day.nunique() < 2:
        st.caption("Baru 1 tanggal tersimpan, grafik tren butuh minimal 2 tanggal.")
    a, b = st.columns(2)
    with a:
        st.caption("Total Sales")
        st.line_chart(ts.groupby(day)["SALES"].sum())
        st.caption("Rata-rata STD")
        st.line_chart(ts.groupby(day)["STD"].mean())
    with b:
        st.caption("Rata-rata SPD")
        st.line_chart(ts.groupby(day)["SPD"].mean())
        st.caption("Rata-rata APC")
        st.line_chart(ts.groupby(day)["APC"].mean())

    # ---- OTD / LATE (dari SLA)
    if mds_sla is not None and not mds_sla.empty:
        sl = mds_sla if not cabang else mds_sla[mds_sla["NAMA_BRANCH"].isin(cabang)]
        sl = apply_search(sl, q)
        st.subheader("OTD dan LATE")
        daily = sl.groupby(sl["TANGGAL"].dt.normalize())[
            ["DELIVERY_ONTIME", "DELIVERY_LATE", "JUMLAH_DELIVERY"]].sum()
        trend = pd.DataFrame({
            "%OTD": daily["DELIVERY_ONTIME"] / daily["JUMLAH_DELIVERY"] * 100,
            "%LATE": daily["DELIVERY_LATE"] / daily["JUMLAH_DELIVERY"] * 100,
        })
        st.caption("Tren harian MDS (%)")
        st.line_chart(trend)
        one = sl[sl["TANGGAL"].dt.normalize() == as_of].copy()
        if not one.empty:
            one["%OTD"] = one["DELIVERY_ONTIME"] / one["JUMLAH_DELIVERY"] * 100
            one["%LATE"] = one["DELIVERY_LATE"] / one["JUMLAH_DELIVERY"] * 100
            st.caption(f"Per toko, {as_of.strftime('%d %b %Y')}")
            tbl = (one[["KD_STORE", "NAMA_STORE", "NAMA_BRANCH", "JUMLAH_DELIVERY", "%OTD", "%LATE"]]
                   .sort_values("%OTD"))
            st.dataframe(styled(tbl, ["JUMLAH_DELIVERY"], ["%OTD", "%LATE"]),
                         use_container_width=True, hide_index=True)

    # ---- snapshot report MDS (SPD, STD, APC, GM, OOS)
    if mds is not None and not mds.empty:
        st.subheader("Snapshot report MDS")
        m = mds if not cabang else mds[mds["NAMA_BRANCH"].isin(cabang)]
        m = apply_search(m, q)
        snap_dates = sorted(m["TANGGAL"].dt.normalize().unique())
        snap_date = st.selectbox("Tanggal snapshot", snap_dates[::-1], key="snap_mds",
                                 format_func=lambda d: pd.Timestamp(d).strftime("%d %b %Y"))
        snap = m[m["TANGGAL"].dt.normalize() == pd.Timestamp(snap_date)].copy()
        pct_cols = [c for c in ["D_GM", "D_OOS", "D_OTD", "D_LATE", "M_GM", "M_OOS", "M_OTD", "M_LATE"]
                    if c in snap.columns]
        for c in pct_cols:
            snap[c] = snap[c] * 100
        int_cols = [c for c in ["D_JHK", "D_SPD", "D_SPD_TAGI", "D_STD", "D_APC", "M_JHK", "M_SALES",
                                "M_SALES_TAGI", "M_SPD", "M_SPD_TAGI", "M_STD", "M_APC"]
                    if c in snap.columns]
        st.dataframe(styled(snap.drop(columns=["TANGGAL"]), int_cols, pct_cols, pct_dec=2),
                     use_container_width=True, hide_index=True)


# ================================================================ HALAMAN: COMPARE TOKO
def page_compare():
    st.title("Compare Toko")
    sumber = st.radio("Sumber data", ["Growth DS", "MDS"], horizontal=True, key="cmp_src")
    base = hist if sumber == "Growth DS" else mds_hist
    if base is None or base.empty:
        st.info("Data belum tersedia untuk sumber ini.")
        return

    sdf = base.drop_duplicates("KD_STORE", keep="last")[["KD_STORE", "NAMA_STORE", "NAMA_BRANCH"]].copy()
    sdf["label"] = (sdf["KD_STORE"].astype(str) + " - " + sdf["NAMA_STORE"].astype(str)
                    + " (" + sdf["NAMA_BRANCH"].astype(str) + ")")
    sdf = sdf.sort_values("label")
    label2code = dict(zip(sdf["label"], sdf["KD_STORE"]))
    short = dict(zip(sdf["KD_STORE"], sdf["KD_STORE"].astype(str) + " " + sdf["NAMA_STORE"].astype(str)))
    metric_opts = {"Sales": "SALES", "SPD": "SPD", "STD": "STD", "APC": "APC",
                   "%GM": "PERCENT_GM", "%OOS OFMB": "PCT_OOS_OFMB", "JHK": "JHK"}

    with st.sidebar:
        st.header("Filter")
        reload_button("reload_cmp")
        chosen = st.multiselect("Pilih toko (2-6)", list(label2code), max_selections=6, key="cmp_stores")
        dates = sorted(base["TANGGAL"].dt.normalize().unique())
        bad = incomplete_dates(base)
        pick, use_all = pick_dates(bad, dates, "all_cmp")
        as_of = st.selectbox("Per tanggal", pick[::-1], key="asof_cmp",
                             format_func=lambda d: pd.Timestamp(d).strftime("%d %b %Y"))
        metric = st.selectbox("Metrik tren", list(metric_opts), key="cmp_metric")

    if len(chosen) < 2:
        st.info("Pilih minimal 2 toko di sidebar (bisa cari dengan mengetik kode atau nama toko).")
        return

    as_of = pd.Timestamp(as_of)
    codes = [label2code[c] for c in chosen]
    sub = base[base["KD_STORE"].isin(codes)].copy()
    sub["Toko"] = sub["KD_STORE"].map(short)

    # ---- tren
    col = metric_opts[metric]
    ts = sub if use_all else sub[~sub["TANGGAL"].dt.normalize().isin(bad)]
    tr = (ts.assign(_d=ts["TANGGAL"].dt.normalize())
          .pivot_table(index="_d", columns="Toko", values=col, aggfunc="mean"))
    if col in ("PERCENT_GM", "PCT_OOS_OFMB"):
        tr = tr * 100
    st.subheader(f"Tren {metric}")
    st.line_chart(tr)

    # ---- tabel perbandingan
    g = compute_growth(sub, as_of)
    month_start = as_of.replace(day=1)
    dd = sub["TANGGAL"].dt.normalize()
    day = sub[dd == as_of][["KD_STORE", "JHK", "SPD", "STD", "APC", "PERCENT_GM", "PCT_OOS_OFMB"]]
    mt = (sub[(dd >= month_start) & (dd <= as_of)]
          .groupby("KD_STORE")[["SPD", "STD", "APC", "PERCENT_GM"]].mean()
          .add_suffix(" MTD").reset_index())
    tab = g.merge(day, on="KD_STORE", how="left").merge(mt, on="KD_STORE", how="left")

    if sumber == "MDS" and mds_sla is not None and not mds_sla.empty:
        sl = mds_sla[mds_sla["KD_STORE"].isin(codes)]
        sd = sl["TANGGAL"].dt.normalize()
        cols_sla = ["DELIVERY_ONTIME", "JUMLAH_DELIVERY"]
        d1 = sl[sd == as_of].groupby("KD_STORE")[cols_sla].sum()
        d2 = sl[(sd >= month_start) & (sd <= as_of)].groupby("KD_STORE")[cols_sla].sum()
        otd = pd.DataFrame({
            "%OTD": d1["DELIVERY_ONTIME"] / d1["JUMLAH_DELIVERY"].replace(0, float("nan")) * 100,
            "%OTD MTD": d2["DELIVERY_ONTIME"] / d2["JUMLAH_DELIVERY"].replace(0, float("nan")) * 100,
        }).reset_index()
        tab = tab.merge(otd, on="KD_STORE", how="left")

    for c in ["Growth DoD", "Growth WoW", "Growth MTD", "PERCENT_GM", "PCT_OOS_OFMB", "PERCENT_GM MTD"]:
        if c in tab.columns:
            tab[c] = tab[c] * 100
    tab = tab.rename(columns={"PERCENT_GM": "%GM", "PCT_OOS_OFMB": "%OOS OFMB", "PERCENT_GM MTD": "%GM MTD"})
    tab["Toko"] = tab["KD_STORE"].map(short)

    order = ["Toko", "NAMA_BRANCH", "Sales", "Growth DoD", "Growth WoW", "MTD", "MTD bulan lalu",
             "Growth MTD", "JHK", "SPD", "STD", "APC", "%GM", "%OOS OFMB",
             "SPD MTD", "STD MTD", "APC MTD", "%GM MTD", "%OTD", "%OTD MTD"]
    tab = tab[[c for c in order if c in tab.columns]]
    pct_cols = ["Growth DoD", "Growth WoW", "Growth MTD", "%GM", "%OOS OFMB", "%GM MTD", "%OTD", "%OTD MTD"]
    int_cols = ["Sales", "MTD", "MTD bulan lalu", "JHK", "SPD", "STD", "APC", "SPD MTD", "STD MTD", "APC MTD"]

    st.subheader(f"Perbandingan per {as_of.strftime('%d %b %Y')}")
    st.dataframe(styled(tab, int_cols, pct_cols), use_container_width=True, hide_index=True)

    # ---- grafik batang antar toko
    a, b = st.columns(2)
    with a:
        st.caption("Sales MTD")
        st.bar_chart(tab.set_index("Toko")["MTD"])
    with b:
        if "SPD MTD" in tab.columns:
            st.caption("Rata-rata SPD MTD")
            st.bar_chart(tab.set_index("Toko")["SPD MTD"])


# ---------------------------------------------------------------- navbar
pages = [
    st.Page(page_ds, title="Growth DS", default=True),
    st.Page(page_mds, title="MDS", url_path="mds"),
    st.Page(page_compare, title="Compare Toko", url_path="compare"),
]
try:
    nav = st.navigation(pages, position="top")   # navbar di atas (Streamlit >= 1.46)
except TypeError:
    nav = st.navigation(pages)                   # Streamlit lama: navigasi di sidebar
nav.run()
