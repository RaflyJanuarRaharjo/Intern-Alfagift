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

## Web App (Streamlit)

Keenam notebook juga tersedia sebagai **satu website**: pilih laporan di sidebar, upload file mentah (satu-satu atau 1 ZIP), klik **Proses**, lalu download hasilnya.

```
app.py                 # tampilan web + navigasi
core/                  # logika tiap laporan (dipindah dari notebook)
  darkstore_ds.py        # Daily Performance DS + MTD DS Performance Delivery
  inventory_report.py    # Inventory, OOS, MAT, Store Performance
  mtd_performance.py     # Report Performance Darkstore MTD
  master_produk.py       # Merge Master Produk + filter TAG
  apo_darkstore.py       # Update APO Darkstore (Excel + PNG)
  common.py              # ekstrak ZIP, recalc LibreOffice
templates/             # template bawaan report Inventory
tests/smoke_test.py    # tes semua laporan dengan data dummy
```

### Menjalankan secara lokal

1. Buat virtual env & install dependensi:
   ```bash
   python -m venv .venv
   .venv\Scripts\activate        # Windows
   # source .venv/bin/activate    # macOS/Linux
   pip install -r requirements.txt
   ```
2. (Opsional) Install [LibreOffice](https://www.libreoffice.org/download/) supaya sel formula di Daily/MTD DS langsung punya nilai. Tanpa LibreOffice laporan tetap jadi; Excel menghitung formulanya saat file dibuka.
3. Jalankan:
   ```bash
   streamlit run app.py
   ```
4. Buka `http://localhost:8501`.

Cek cepat semua laporan dengan data dummy: `python tests/smoke_test.py`.

> **Keamanan data:** jangan deploy ke hosting publik (mis. Streamlit Community Cloud) tanpa proteksi login, karena file yang diproses berisi data internal. Jalankan lokal atau di server internal. File upload hanya disimpan sementara selama proses lalu dihapus.

## Cara Menjalankan (Notebook)

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
