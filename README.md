# ROC–ARAS Flask App

Fitur:
1) Login & Register
2) Upload data Excel (raw) -> proses otomatis (cleaning, konversi) -> simpan CSV terproses
3) CRUD Kriteria dengan bobot otomatis Rank Order Centroid (ROC) berdasarkan `display_order`
4) CRUD Subkriteria (opsional, untuk dokumentasi aturan)
5) Lihat perhitungan ARAS mulai dari matriks A dan matriks ternormalisasi
6) Hasil perangkingan siswa (ROC–ARAS), bisa difilter berdasarkan `Jurusan` & `Periode` jika kolom ada, dan bisa diunduh CSV

## Menjalankan

```bash
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate
pip install -r requirements.txt
python run.py
```

Buka: http://127.0.0.1:5000
- Register user baru
- Upload Excel (.xlsx). Kolom yang didukung (disarankan): 
  - `Nama Siswa`, `Nilai rata rata Raport`, `Nilai Sikap`, `Nilai Ketidakhadiran`,
    `Nilai Estrakulikuler`, `Nilai Prestasi`, `Nilai Prakerin`, opsi: `Jurusan`, `Periode`.
- Kelola kriteria (urutan menentukan bobot ROC). 
- Lihat Matriks & Ranking.

## Catatan
- Bobot ROC dihitung runtut dari `display_order` (1 = paling penting).
- ARAS membentuk A0 (solusi ideal) dan menggunakan normalisasi sesuai tipe kriteria (benefit/cost).
- Download hasil tersedia pada halaman Ranking (CSV).
