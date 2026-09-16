# Intern Alfagift — Report Automation

Kumpulan notebook Python (Google Colab) untuk mengotomasi pembuatan laporan operasional **Darkstore Alfagift** selama program magang. Tujuannya sederhana: cukup upload file data mentah, notebook akan menghasilkan laporan Excel yang sudah terformat.

## Struktur Repository

```
Intern-Alfagift/
├── notebooks/
│   ├── daily/
│   │   ├── report_daily_performance_ds.ipynb
│   │   └── report_daily_inventory.ipynb
│   ├── mtd/
│   │   ├── mtd_ds_performance_delivery.ipynb
│   │   └── mtd_firstman.ipynb
│   └── master-data/
│       ├── merge_master_produk_filter_tag_dgsw6.ipynb
│       └── update_apo_darkstore_bitmap.ipynb
├── data/          # data mentah (tidak di-commit)
├── output/        # hasil laporan (tidak di-commit)
├── requirements.txt
├── .gitignore
└── README.md
```

## Daftar Notebook

| Kategori | Notebook | Fungsi |
|---|---|---|
| Daily | `report_daily_performance_ds.ipynb` | Membuat *Report Daily Performance Darkstore* dari Detail Data (CSV) dan OOS By Toko (XLSX) |
| Daily | `report_daily_inventory.ipynb` | Membuat laporan harian Inventory, OOS, MAT, dan Store Performance |
| MTD | `mtd_ds_performance_delivery.ipynb` | Rekap *month-to-date* performa darkstore dan delivery (heatmap, urutan % on-time) |
| MTD | `mtd_firstman.ipynb` | Rekap *month-to-date* FirstMan |
| Master Data | `merge_master_produk_filter_tag_dgsw6.ipynb` | Menggabungkan master produk dan memfilter tag DGSW6 |
| Master Data | `update_apo_darkstore_bitmap.ipynb` | Memperbarui data bitmap APO Darkstore |

> Silakan sesuaikan deskripsi di atas bila ada yang kurang tepat.

## Cara Menjalankan

1. Buka notebook di Google Colab (klik file → tombol **Open in Colab**, atau upload manual).
2. Jalankan sel instalasi/import di bagian atas.
3. Upload file data mentah saat diminta.
4. Jalankan semua sel (`Runtime → Run all`).
5. File laporan `.xlsx` akan otomatis terunduh.

Untuk menjalankan secara lokal:

```bash
pip install -r requirements.txt
jupyter notebook
```

## Catatan Keamanan Data

Data mentah dan hasil laporan berisi data internal perusahaan, jadi **tidak di-commit** ke repository ini. Output sel notebook juga dibersihkan sebelum commit.

## Author

**Rafly Januar Raharjo** — Intern Alfagift
