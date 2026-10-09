"""Dashboard Performance Darkstore (read-only).

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
        return None, None
    return storage.load_history(backend), storage.load_oos(backend)


hist, oos = load()
if hist is None:
    st.error("Penyimpanan belum diatur (secrets `gcp_service_account` dan `storage.spreadsheet_id`).")
    st.stop()
if hist.empty:
    st.info("Belum ada data tersimpan. Jalankan app olah data dulu.")
    st.stop()

st.title("Dashboard Performance Darkstore")

# ---------------------------------------------------------------- filter
with st.sidebar:
    st.header("Filter")
    if st.button("Muat ulang data"):
        st.cache_data.clear()
        st.rerun()
    cabang = st.multiselect("Cabang", sorted(hist["NAMA_BRANCH"].dropna().unique()))
    dates = sorted(hist["TANGGAL"].dt.normalize().unique())
    bad_dates = incomplete_dates(hist)
    use_all = st.checkbox("Sertakan tanggal yang datanya belum lengkap", value=False)
    pick = dates if (use_all or not bad_dates) else [d for d in dates if pd.Timestamp(d) not in set(bad_dates)]
    pick = pick or dates
    as_of = st.selectbox("Per tanggal", pick[::-1],
                         format_func=lambda d: pd.Timestamp(d).strftime("%d %b %Y"))

view = hist if not cabang else hist[hist["NAMA_BRANCH"].isin(cabang)]
as_of = pd.Timestamp(as_of)
g = compute_growth(view, as_of)

if bad_dates and not use_all:
    st.warning("Tanggal berikut disembunyikan karena datanya tampak belum lengkap "
               "(total Sales jauh di bawah hari-hari sebelumnya): "
               + ", ".join(pd.Timestamp(d).strftime("%d %b") for d in bad_dates)
               + ". Centang 'Sertakan tanggal...' di sidebar untuk menampilkannya.")

# ---------------------------------------------------------------- KPI
month_start = as_of.replace(day=1)
mtd = view[(view["TANGGAL"].dt.normalize() >= month_start) & (view["TANGGAL"].dt.normalize() <= as_of)]

c1, c2, c3, c4 = st.columns(4)
c1.metric(f"Sales {as_of.strftime('%d %b')}", f"{g['Sales'].sum():,.0f}")
c2.metric("Sales MTD", f"{mtd['SALES'].sum():,.0f}")
c3.metric("Rata-rata SPD (MTD)", f"{mtd['SPD'].mean():,.0f}")
c4.metric("Rata-rata %GM (MTD)", f"{mtd['PERCENT_GM'].mean() * 100:,.2f}%")

# ---------------------------------------------------------------- tren
st.subheader("Tren Sales harian")
trend_src = view if use_all else view[~view["TANGGAL"].dt.normalize().isin(bad_dates)]
trend = trend_src.groupby(trend_src["TANGGAL"].dt.normalize())["SALES"].sum()
st.line_chart(trend)

# ---------------------------------------------------------------- tabel growth
st.subheader("Performa per toko")
if not has_previous_month(hist, as_of):
    st.caption("Growth MTD vs bulan lalu kosong: riwayat bulan sebelumnya belum tersimpan.")

show = g.copy()
for col in ["Growth DoD", "Growth WoW", "Growth MTD"]:
    show[col] = show[col] * 100  # tampil dalam persen
st.dataframe(
    show[["KD_STORE", "NAMA_STORE", "NAMA_BRANCH", "Sales", "Growth DoD", "Growth WoW",
          "MTD", "MTD bulan lalu", "Growth MTD"]],
    use_container_width=True, hide_index=True,
    column_config={
        "Sales": st.column_config.NumberColumn(format="%,d"),
        "MTD": st.column_config.NumberColumn(format="%,d"),
        "MTD bulan lalu": st.column_config.NumberColumn(format="%,d"),
        "Growth DoD": st.column_config.NumberColumn(format="%.1f%%"),
        "Growth WoW": st.column_config.NumberColumn(format="%.1f%%"),
        "Growth MTD": st.column_config.NumberColumn(format="%.1f%%"),
    },
)

# ---------------------------------------------------------------- top / bottom
left, right = st.columns(2)
with left:
    st.subheader("5 toko Sales MTD tertinggi")
    st.bar_chart(g.nlargest(5, "MTD").set_index("NAMA_STORE")["MTD"])
with right:
    st.subheader("5 toko Sales MTD terendah")
    st.bar_chart(g.nsmallest(5, "MTD").set_index("NAMA_STORE")["MTD"])

# ---------------------------------------------------------------- OOS
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


# ======================================================================
# Growth MDS (tambahan)
# ======================================================================
@st.cache_data(ttl=300, show_spinner="Memuat data MDS...")
def load_mds_data():
    backend = storage.get_backend_from_streamlit()
    if backend is None:
        return None
    return storage.load_mds(backend)


st.divider()
st.header("Growth MDS (Mini Darkstore)")

mds = load_mds_data()
if mds is None or mds.empty:
    st.info("Data MDS belum tersimpan. Jalankan app olah data (report MDS) dulu.")
else:
    if cabang:
        mds = mds[mds["NAMA_BRANCH"].isin(cabang)]
    mds_codes = set(mds["KD_STORE"].unique())

    # --- growth DoD / WoW / MTD dari riwayat, khusus toko MDS
    mds_hist = view[view["KD_STORE"].isin(mds_codes)]
    if mds_hist.empty:
        st.info("Riwayat toko MDS belum ada di data history.")
    else:
        gm = compute_growth(mds_hist, as_of)

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Jumlah toko MDS", f"{len(gm):,}")
        k2.metric(f"Sales MDS {as_of.strftime('%d %b')}", f"{gm['Sales'].sum():,.0f}")
        k3.metric("Sales MDS MTD", f"{gm['MTD'].sum():,.0f}")
        prev = gm["MTD bulan lalu"].sum()
        k4.metric("Growth MTD MDS",
                  f"{(gm['MTD'].sum() / prev - 1) * 100:,.1f}%" if prev else "-")

        st.subheader("Performa MDS per toko")
        shm = gm.copy()
        for col in ["Growth DoD", "Growth WoW", "Growth MTD"]:
            shm[col] = shm[col] * 100
        st.dataframe(
            shm[["KD_STORE", "NAMA_STORE", "NAMA_BRANCH", "Sales", "Growth DoD",
                 "Growth WoW", "MTD", "MTD bulan lalu", "Growth MTD"]],
            use_container_width=True, hide_index=True,
            column_config={
                "Sales": st.column_config.NumberColumn(format="%.0f"),
                "MTD": st.column_config.NumberColumn(format="%.0f"),
                "MTD bulan lalu": st.column_config.NumberColumn(format="%.0f"),
                "Growth DoD": st.column_config.NumberColumn(format="%.1f%%"),
                "Growth WoW": st.column_config.NumberColumn(format="%.1f%%"),
                "Growth MTD": st.column_config.NumberColumn(format="%.1f%%"),
            },
        )

    # --- snapshot report MDS (SPD, STD, APC, GM, OOS, OTD, LATE)
    st.subheader("Snapshot report MDS")
    mds_dates = sorted(mds["TANGGAL"].dt.normalize().unique())
    snap_date = st.selectbox("Tanggal snapshot MDS", mds_dates[::-1],
                             format_func=lambda d: pd.Timestamp(d).strftime("%d %b %Y"))
    snap = mds[mds["TANGGAL"].dt.normalize() == pd.Timestamp(snap_date)].copy()
    pct_cols = [c for c in ["D_GM", "D_OOS", "D_OTD", "D_LATE",
                            "M_GM", "M_OOS", "M_OTD", "M_LATE"] if c in snap.columns]
    for c in pct_cols:
        snap[c] = snap[c] * 100
    st.dataframe(
        snap.drop(columns=["TANGGAL"]), use_container_width=True, hide_index=True,
        column_config={c: st.column_config.NumberColumn(format="%.2f%%") for c in pct_cols},
    )

    # --- tren MDS dari snapshot tersimpan
    st.subheader("Tren MDS (rata-rata antar toko)")
    trend_cols = [c for c in ["D_SPD", "D_STD", "D_APC"] if c in mds.columns]
    if trend_cols:
        tr = mds.groupby(mds["TANGGAL"].dt.normalize())[trend_cols].mean()
        st.line_chart(tr)
    if "D_OTD" in mds.columns and mds["D_OTD"].notna().any():
        otd = mds.groupby(mds["TANGGAL"].dt.normalize())[["D_OTD", "D_LATE"]].mean() * 100
        st.caption("OTD dan LATE (%)")
        st.line_chart(otd)
