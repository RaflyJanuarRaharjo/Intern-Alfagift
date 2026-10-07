"""Merge Master Produk + filter TAG.

Sumber: notebooks/master-data/merge_master_produk_filter_tag_dgsw6.ipynb.
Di notebook file dibaca dari Google Drive; di web file di-upload langsung
(boleh banyak file .xls sekaligus atau 1 ZIP).
"""
import os

import pandas as pd

from .common import Log

DEFAULT_TAGS = ['D', 'G', 'S', 'W', '6']


def read_master_file(path):
    """File .xls dari sistem sebenarnya teks tab-separated, 2 baris judul di atas header."""
    return pd.read_csv(path, sep='\t', skiprows=2)


def run(paths, output_dir, tags=None, log=None):
    log = log or Log()
    tags = [str(t).strip() for t in (tags or DEFAULT_TAGS) if str(t).strip()]
    files = [p for p in paths if p.lower().endswith(('.xls', '.txt', '.tsv', '.csv'))]
    log(f"Jumlah file ditemukan: {len(files)}")

    frames = []
    for p in files:
        try:
            frames.append(read_master_file(p))
        except Exception as e:  # noqa: BLE001
            log(f"Gagal membaca file {os.path.basename(p)}: {e}")
    if not frames:
        raise ValueError("Tidak ada data yang berhasil dibaca.")

    df_merge = pd.concat(frames, ignore_index=True)
    log(f"Total baris gabungan: {len(df_merge)}, kolom: {len(df_merge.columns)}")
    for col in ('TAG', 'GROUP CODE'):
        if col not in df_merge.columns:
            raise KeyError(f"Kolom '{col}' tidak ada. Kolom yang tersedia: {list(df_merge.columns)}")

    df_merge['TAG'] = df_merge['TAG'].astype(str).str.strip()
    df_filtered = df_merge[df_merge['TAG'].isin(tags)]
    log(f"Baris setelah filter TAG {tags}: {len(df_filtered)}")

    groups_without_tags = sorted(set(df_merge['GROUP CODE']) - set(df_filtered['GROUP CODE']), key=str)
    df_missing = pd.DataFrame([
        {'GROUP CODE': g,
         'TAG YANG TERSEDIA': ", ".join(map(str, df_merge.loc[df_merge['GROUP CODE'] == g, 'TAG'].unique()))}
        for g in groups_without_tags
    ], columns=['GROUP CODE', 'TAG YANG TERSEDIA'])
    if len(df_missing):
        log(f"[Info] {len(df_missing)} group code tidak punya tag {tags} (lihat sheet 'Group Tanpa Tag').")

    out = os.path.join(output_dir, 'data_produk_filtered.xlsx')
    with pd.ExcelWriter(out, engine='openpyxl') as xw:
        df_filtered.to_excel(xw, sheet_name='Data Filtered', index=False)
        df_missing.to_excel(xw, sheet_name='Group Tanpa Tag', index=False)
    return out, df_missing, []
