"""Dashboard Performance Darkstore (read-only), dengan navbar: Growth DS | MDS.

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
        return None, None, None
    return storage.load_history(backend), storage.load_oos(backend), storage.load_mds(backend)


hist, oos, mds = load()
if hist is None:
    st.error("Penyimpanan belum diatur (secrets `gcp_service_account` dan `storage.spreadsheet_id`).")
    st.stop()
if hist.empty:
    st.info("Belum ada data tersimpan. Jalankan app olah data dulu.")
    st.stop()


def reload_button(key):
    if st.button("Muat ulang data", key=key):
        st.cache_data.clear()
        st.rerun()


def pick_dates(bad_dates, dates, key):
    use_all = st.checkbox("Sertakan tanggal yang datanya belum lengkap", value=False, key=key)
    pick = dates if (use_all or not bad_dates) else [d for d in dates if pd.Timestamp(d) not in set(bad_dates)]
    return (pick or dates), use_all


PCT_FMT = st.column_config.NumberColumn(format="%.1f%%")
INT_FMT = st.column_config.NumberColumn(format="%.0f")


# ================================================================ HALAMAN: GROWTH DS
def page_ds():
    st.title("Dashboard Performance Darkstore")

    with st.sidebar:
        st.header("Filter")
        reload_button("reload_ds")
        cabang = st.multiselect("Cabang", sorted(hist["NAMA_BRANCH"].dropna().unique()), key="cab_ds")
        dates = sorted(hist["TANGGAL"].dt.normalize().unique())
        bad_dates = incomplete_dates(hist)
        pick, use_all = pick_dates(bad_dates, dates, "all_ds")
        as_of = st.selectbox("Per tanggal", pick[::-1], key="asof_ds",
                             format_func=lambda d: pd.Timestamp(d).strftime("%d %b %Y"))

    view = hist if not cabang else hist[hist["NAMA_BRANCH"].isin(cabang)]
    as_of = pd.Timestamp(as_of)
    g = compute_growth(view, as_of)

    if bad_dates and not use_all:
        st.warning("Tanggal berikut disembunyikan karena datanya tampak belum lengkap "
                   "(total Sales jauh di bawah hari-hari sebelumnya): "
                   + ", ".join(pd.Timestamp(d).strftime("%d %b") for d in bad_dates)
                   + ". Centang 'Sertakan tanggal...' di sidebar untuk menampilkannya.")

    # KPI
    month_start = as_of.replace(day=1)
    mtd = view[(view["TANGGAL"].dt.normalize() >= month_start) & (view["TANGGAL"].dt.normalize() <= as_of)]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric(f"Sales {as_of.strftime('%d %b')}", f"{g['Sales'].sum():,.0f}")
    c2.metric("Sales MTD", f"{mtd['SALES'].sum():,.0f}")
    c3.metric("Rata-rata SPD (MTD)", f"{mtd['SPD'].mean():,.0f}")
    c4.metric("Rata-rata %GM (MTD)", f"{mtd['PERCENT_GM'].mean() * 100:,.2f}%")

    # tren
    st.subheader("Tren Sales harian")
    trend_src = view if use_all else view[~view["TANGGAL"].dt.normalize().isin(bad_dates)]
    trend = trend_src.groupby(trend_src["TANGGAL"].dt.normalize())["SALES"].sum()
    st.line_chart(trend)

    # tabel growth
    st.subheader("Performa per toko")
    if not has_previous_month(hist, as_of):
        st.caption("Growth MTD vs bulan lalu kosong: riwayat bulan sebelumnya belum tersimpan.")
    show = g.copy()
    for col in ["Growth DoD", "Growth WoW", "Growth MTD"]:
        show[col] = show[col] * 100
    st.dataframe(
        show[["KD_STORE", "NAMA_STORE", "NAMA_BRANCH", "Sales", "Growth DoD", "Growth WoW",
              "MTD", "MTD bulan lalu", "Growth MTD"]],
        use_container_width=True, hide_index=True,
        column_config={"Sales": INT_FMT, "MTD": INT_FMT, "MTD bulan lalu": INT_FMT,
                       "Growth DoD": PCT_FMT, "Growth WoW": PCT_FMT, "Growth MTD": PCT_FMT},
    )

    # top / bottom
    left, right = st.columns(2)
    with left:
        st.subheader("5 toko Sales MTD tertinggi")
        st.bar_chart(g.nlargest(5, "MTD").set_index("NAMA_STORE")["MTD"])
    with right:
        st.subheader("5 toko Sales MTD terendah")
        st.bar_chart(g.nsmallest(5, "MTD").set_index("NAMA_STORE")["MTD"])

    # OOS
    if oos is not None and not oos.empty:
        st.subheader("OOS periode terbaru")
        latest = oos["PERIODE_END"].max()
        o = oos[oos["PERIODE_END"] == latest].merge(
            hist.drop_duplicates("KD_STORE", keep="last")[["KD_STORE", "NAMA_STORE", "NAMA_BRANCH"]],
            on="KD_STORE", how="left")
        if cabang:
            o = o[o["NAMA_BRANCH"].isin(cabang)]
        st.caption(f"Periode s.d. {latest}")
        for col in ["% OOS OFMB", "% OOS TAG I", "% OOS TAG K"]:
            if col in o.columns:
                o[col] = o[col] * 100
        st.dataframe(o.sort_values("% OOS OFMB", ascending=False), use_container_width=True, hide_index=True,
                     column_config={c: st.column_config.NumberColumn(format="%.2f%%")
                                    for c in ["% OOS OFMB", "% OOS TAG I", "% OOS TAG K"] if c in o.columns})


# ================================================================ HALAMAN: MDS
def page_mds():
    st.title("Dashboard MDS (Mini Darkstore)")

    if mds is None or mds.empty:
        st.info("Data MDS belum tersimpan. Jalankan app olah data (report MDS) dulu.")
        return

    with st.sidebar:
        st.header("Filter")
        reload_button("reload_mds")
        cabang = st.multiselect("Cabang", sorted(mds["NAMA_BRANCH"].dropna().unique()), key="cab_mds")

    m = mds if not cabang else mds[mds["NAMA_BRANCH"].isin(cabang)]
    codes = set(m["KD_STORE"].unique())
    mh = hist[hist["KD_STORE"].isin(codes)]
    if mh.empty:
        st.info("Riwayat toko MDS belum ada di data history (jalankan report Darkstore untuk tanggal yang sama).")
        return

    with st.sidebar:
        dates = sorted(mh["TANGGAL"].dt.normalize().unique())
        bad_dates = incomplete_dates(hist)
        pick, use_all = pick_dates(bad_dates, dates, "all_mds")
        as_of = st.selectbox("Per tanggal", pick[::-1], key="asof_mds",
                             format_func=lambda d: pd.Timestamp(d).strftime("%d %b %Y"))
    as_of = pd.Timestamp(as_of)

    # ---- KPI + growth (dari history, khusus toko MDS)
    gm = compute_growth(mh, as_of)
    prev = gm["MTD bulan lalu"].sum()
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Jumlah toko MDS", f"{len(gm):,}")
    k2.metric(f"Sales MDS {as_of.strftime('%d %b')}", f"{gm['Sales'].sum():,.0f}")
    k3.metric("Sales MDS MTD", f"{gm['MTD'].sum():,.0f}")
    k4.metric("Growth MTD MDS", f"{(gm['MTD'].sum() / prev - 1) * 100:,.1f}%" if prev else "-")

    st.subheader("Performa MDS per toko")
    show = gm.copy()
    for col in ["Growth DoD", "Growth WoW", "Growth MTD"]:
        show[col] = show[col] * 100
    st.dataframe(
        show[["KD_STORE", "NAMA_STORE", "NAMA_BRANCH", "Sales", "Growth DoD", "Growth WoW",
              "MTD", "MTD bulan lalu", "Growth MTD"]],
        use_container_width=True, hide_index=True,
        column_config={"Sales": INT_FMT, "MTD": INT_FMT, "MTD bulan lalu": INT_FMT,
                       "Growth DoD": PCT_FMT, "Growth WoW": PCT_FMT, "Growth MTD": PCT_FMT},
    )

    # ---- tren harian (dari history, semua tanggal tersedia)
    st.subheader("Tren harian MDS")
    ts = mh if use_all else mh[~mh["TANGGAL"].dt.normalize().isin(bad_dates)]
    day = ts["TANGGAL"].dt.normalize()
    if day.nunique() < 2:
        st.caption("Baru 1 tanggal di data history, grafik tren butuh minimal 2 tanggal.")
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

    # ---- snapshot report MDS (termasuk OTD / LATE)
    st.subheader("Snapshot report MDS")
    snap_dates = sorted(m["TANGGAL"].dt.normalize().unique())
    snap_date = st.selectbox("Tanggal snapshot", snap_dates[::-1], key="snap_mds",
                             format_func=lambda d: pd.Timestamp(d).strftime("%d %b %Y"))
    snap = m[m["TANGGAL"].dt.normalize() == pd.Timestamp(snap_date)].copy()
    pct_cols = [c for c in ["D_GM", "D_OOS", "D_OTD", "D_LATE", "M_GM", "M_OOS", "M_OTD", "M_LATE"]
                if c in snap.columns]
    for c in pct_cols:
        snap[c] = snap[c] * 100
    st.dataframe(snap.drop(columns=["TANGGAL"]), use_container_width=True, hide_index=True,
                 column_config={c: st.column_config.NumberColumn(format="%.2f%%") for c in pct_cols})

    if "D_OTD" in snap.columns and snap["D_OTD"].notna().any():
        st.caption(f"OTD dan LATE per toko (%), {pd.Timestamp(snap_date).strftime('%d %b %Y')}")
        st.bar_chart(snap.set_index("NAMA_STORE")[["D_OTD", "D_LATE"]])
    if len(snap_dates) > 1 and "D_OTD" in m.columns:
        st.caption("Tren OTD dan LATE (rata-rata antar toko, %)")
        st.line_chart(m.groupby(m["TANGGAL"].dt.normalize())[["D_OTD", "D_LATE"]].mean() * 100)
    elif len(snap_dates) == 1:
        st.caption("Tren OTD/LATE muncul setelah report MDS diproses untuk lebih dari 1 tanggal.")


# ---------------------------------------------------------------- navbar
pages = [
    st.Page(page_ds, title="Growth DS", default=True),
    st.Page(page_mds, title="MDS", url_path="mds"),
]
try:
    nav = st.navigation(pages, position="top")   # navbar di atas (Streamlit baru)
except TypeError:
    nav = st.navigation(pages)                   # Streamlit lama: navigasi di sidebar
nav.run()
