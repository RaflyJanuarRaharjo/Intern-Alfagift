"""Hitung growth Sales per toko dari riwayat (DoD, WoW, MTD vs bulan lalu).

Growth = (sekarang - pembanding) / pembanding.
Kalau pembanding 0 / kosong -> hasil dikosongkan (NaN), bukan dipaksa jadi angka.
MTD vs bulan lalu dibandingkan apple-to-apple: tanggal 1..N bulan ini vs tanggal 1..N bulan lalu.
"""
import calendar

import numpy as np
import pandas as pd


def _growth(cur, prev):
    cur = pd.to_numeric(cur, errors="coerce")
    prev = pd.to_numeric(prev, errors="coerce")
    return np.where(prev > 0, cur / prev - 1, np.nan)


def _sales_on(hist, day):
    return hist.loc[hist["TANGGAL"].dt.normalize() == pd.Timestamp(day).normalize()] \
               .groupby("KD_STORE")["SALES"].sum()


def _sales_between(hist, start, end):
    m = hist["TANGGAL"].dt.normalize().between(pd.Timestamp(start), pd.Timestamp(end))
    return hist.loc[m].groupby("KD_STORE")["SALES"].sum()


def compute_growth(hist, as_of=None):
    """Return DataFrame per toko berisi Sales as_of, DoD, WoW, MTD, MTD bulan lalu & growth-nya."""
    if hist.empty:
        return pd.DataFrame()
    as_of = pd.Timestamp(as_of or hist["TANGGAL"].max()).normalize()

    master = (hist.sort_values("TANGGAL")
              .drop_duplicates("KD_STORE", keep="last")[["KD_STORE", "NAMA_STORE", "NAMA_BRANCH"]]
              .set_index("KD_STORE"))
    out = master.copy()

    cur = _sales_on(hist, as_of)
    out["Sales"] = cur
    out["Sales kemarin"] = _sales_on(hist, as_of - pd.Timedelta(days=1))
    out["Sales 7 hari lalu"] = _sales_on(hist, as_of - pd.Timedelta(days=7))
    out["Growth DoD"] = _growth(out["Sales"], out["Sales kemarin"])
    out["Growth WoW"] = _growth(out["Sales"], out["Sales 7 hari lalu"])

    # MTD vs bulan lalu (rentang tanggal sama)
    month_start = as_of.replace(day=1)
    prev_end_of_month = month_start - pd.Timedelta(days=1)
    prev_start = prev_end_of_month.replace(day=1)
    day = min(as_of.day, calendar.monthrange(prev_start.year, prev_start.month)[1])
    prev_end = prev_start.replace(day=day)

    out["MTD"] = _sales_between(hist, month_start, as_of)
    prev_mtd = _sales_between(hist, prev_start, prev_end)
    out["MTD bulan lalu"] = prev_mtd
    out["Growth MTD"] = _growth(out["MTD"], out["MTD bulan lalu"])

    return out.reset_index().sort_values("MTD", ascending=False, na_position="last").reset_index(drop=True)


def has_previous_month(hist, as_of=None):
    """True kalau riwayat memuat bulan sebelum bulan as_of (supaya Growth MTD bermakna)."""
    as_of = pd.Timestamp(as_of or hist["TANGGAL"].max())
    return bool((hist["TANGGAL"] < as_of.replace(day=1)).any())
