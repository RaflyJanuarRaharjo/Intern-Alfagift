"""Hitung growth Sales dan SPD per toko dari riwayat (DoD, WoW, MTD vs bulan lalu).

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


def _spd_on(hist, day):
    m = hist["TANGGAL"].dt.normalize() == pd.Timestamp(day).normalize()
    return hist.loc[m].groupby("KD_STORE")["SPD"].mean()


def _spd_between(hist, start, end):
    m = hist["TANGGAL"].dt.normalize().between(pd.Timestamp(start), pd.Timestamp(end))
    return hist.loc[m].groupby("KD_STORE")["SPD"].mean()


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

    # ---- SPD: DoD / WoW pakai SPD harian; MTD pakai rata-rata SPD harian (sama dengan kartu "Rata-rata SPD (MTD)")
    if "SPD" in hist.columns:
        out["SPD"] = _spd_on(hist, as_of)
        out["SPD kemarin"] = _spd_on(hist, as_of - pd.Timedelta(days=1))
        out["SPD 7 hari lalu"] = _spd_on(hist, as_of - pd.Timedelta(days=7))
        out["Growth SPD DoD"] = _growth(out["SPD"], out["SPD kemarin"])
        out["Growth SPD WoW"] = _growth(out["SPD"], out["SPD 7 hari lalu"])
        out["SPD MTD"] = _spd_between(hist, month_start, as_of)
        out["SPD MTD bulan lalu"] = _spd_between(hist, prev_start, prev_end)
        out["Growth SPD MTD"] = _growth(out["SPD MTD"], out["SPD MTD bulan lalu"])

    return out.reset_index().sort_values("MTD", ascending=False, na_position="last").reset_index(drop=True)


def has_previous_month(hist, as_of=None):
    """True kalau riwayat memuat bulan sebelum bulan as_of (supaya Growth MTD bermakna)."""
    as_of = pd.Timestamp(as_of or hist["TANGGAL"].max())
    return bool((hist["TANGGAL"] < as_of.replace(day=1)).any())


def incomplete_dates(hist, threshold=0.5, window=7, min_history=3):
    """Deteksi tanggal yang datanya kemungkinan BELUM LENGKAP.

    Aturan: total Sales semua toko hari itu < threshold x median total Sales hari-hari
    lengkap sebelumnya (maksimal `window` hari terakhir). Hari pertama (belum ada
    pembanding) dianggap lengkap. Return: list Timestamp, urut naik.
    """
    daily = hist.groupby(hist["TANGGAL"].dt.normalize())["SALES"].sum().sort_index()
    complete_vals, bad = [], []
    for d, v in daily.items():
        ref = complete_vals[-window:]
        if len(ref) >= min_history and v < threshold * float(np.median(ref)):
            bad.append(d)
        else:
            complete_vals.append(v)
    return bad


def last_complete_date(hist, **kw):
    """Tanggal terakhir yang datanya lengkap (None kalau tidak ada)."""
    bad = set(incomplete_dates(hist, **kw))
    days = [d for d in sorted(hist["TANGGAL"].dt.normalize().unique()) if pd.Timestamp(d) not in bad]
    return pd.Timestamp(days[-1]) if days else None


def compute_spd_range(hist, start, end, pembanding="sebelumnya"):
    """Growth rata-rata SPD per toko untuk rentang tanggal [start, end].

    pembanding:
      "sebelumnya"  -> periode tepat sebelum start dengan jumlah hari sama
      "bulan_lalu"  -> rentang tanggal yang sama di bulan sebelumnya
    Rata-rata SPD = rata-rata SPD harian toko pada rentang itu.
    Pembanding kosong / 0 -> growth NaN.
    """
    if hist.empty or "SPD" not in hist.columns:
        return pd.DataFrame()
    start, end = pd.Timestamp(start).normalize(), pd.Timestamp(end).normalize()
    if pembanding == "bulan_lalu":
        p_start, p_end = start - pd.DateOffset(months=1), end - pd.DateOffset(months=1)
    else:
        n = (end - start).days + 1
        p_end = start - pd.Timedelta(days=1)
        p_start = p_end - pd.Timedelta(days=n - 1)

    master = (hist.sort_values("TANGGAL")
              .drop_duplicates("KD_STORE", keep="last")[["KD_STORE", "NAMA_STORE", "NAMA_BRANCH"]]
              .set_index("KD_STORE"))
    out = master.copy()
    out["SPD rentang"] = _spd_between(hist, start, end)
    out["SPD pembanding"] = _spd_between(hist, p_start, p_end)
    out["Growth SPD rentang"] = _growth(out["SPD rentang"], out["SPD pembanding"])
    out = out.dropna(subset=["SPD rentang"])
    return (out.reset_index().sort_values("SPD rentang", ascending=False, na_position="last")
               .reset_index(drop=True)), (p_start, p_end)


def otd_growth_summary(sla, as_of, flat_pp=0.5):
    """Ringkasan growth %OTD (selisih dalam poin persentase) dari tabel SLA per toko per tanggal.

    %OTD gabungan = sum(ONTIME) / sum(JUMLAH_DELIVERY) (bukan rata-rata %OTD tiap toko).
    Dibandingkan: hari ini vs H-1, dan MTD vs MTD bulan lalu (rentang tanggal sama).
    """
    as_of = pd.Timestamp(as_of).normalize()
    cols = ["DELIVERY_ONTIME", "JUMLAH_DELIVERY"]
    d = sla["TANGGAL"].dt.normalize()

    month_start = as_of.replace(day=1)
    prev_start = (month_start - pd.Timedelta(days=1)).replace(day=1)
    day = min(as_of.day, calendar.monthrange(prev_start.year, prev_start.month)[1])
    prev_end = prev_start.replace(day=day)
    kemarin = as_of - pd.Timedelta(days=1)
    periods = {"hari": (as_of, as_of), "kemarin": (kemarin, kemarin),
               "mtd": (month_start, as_of), "mtd_lalu": (prev_start, prev_end)}

    g = {k: sla[(d >= a) & (d <= b)].groupby("KD_STORE")[cols].sum() for k, (a, b) in periods.items()}

    def total(x):
        den = x["JUMLAH_DELIVERY"].sum()
        return x["DELIVERY_ONTIME"].sum() / den * 100 if den > 0 else float("nan")

    def pct(x):
        return x["DELIVERY_ONTIME"] / x["JUMLAH_DELIVERY"].where(x["JUMLAH_DELIVERY"] > 0) * 100

    names = sla.drop_duplicates("KD_STORE", keep="last").set_index("KD_STORE")["NAMA_STORE"]

    def per_store(cur, prev):
        df = pd.DataFrame({"sekarang": pct(g[cur]), "pembanding": pct(g[prev])}).dropna()
        df["selisih"] = df["sekarang"] - df["pembanding"]
        df["NAMA_STORE"] = names.reindex(df.index)
        return df.reset_index().sort_values("selisih", ascending=False).reset_index(drop=True)

    return {
        "as_of": as_of, "kemarin": kemarin, "prev_start": prev_start, "prev_end": prev_end,
        "overall": {k: total(v) for k, v in g.items()},
        "toko_dod": per_store("hari", "kemarin"),
        "toko_mtd": per_store("mtd", "mtd_lalu"),
        "flat_pp": flat_pp,
    }
