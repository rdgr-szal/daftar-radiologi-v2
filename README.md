# Sistem Input Buku Daftar Radiologi & Penjanaan Reten PHRIS V2 (PER.SS-RA 101 Compliance)

Aplikasi berasaskan desktop & local web (localhost offline) yang direka khas untuk Unit Radiologi Klinik Kesihatan, Kementerian Kesihatan Malaysia (KKM). Sistem ini membantu staf Radiologi mendaftarkan pesakit secara digital, memantau analitik operasi di Dashboard, dan menjana reten statistik dalam format PHRIS secara automatik daripada pangkalan data fail Excel rasmi tanpa memerlukan sambungan internet.

---

## 1. Ciri-Ciri Utama Sistem V2

- **Pematuhan Penuh PER.SS-RA 101**: Format pendaftaran rasmi KKM dengan sokongan auto-parsing nombor MyKad/Pasport (Umur, Jantina, Warganegara).
- **Format Tarikh Seragam DD/MM/YYYY**: Penyeragaman format tarikh di seluruh paparan borang, senarai pesakit, kalendar penapis, dan eksport cetakan.
- **Katalog SMRP & Custom Orderables**: Taksonomi rasmi pemeriksaan KKM dengan kebolehan menambah Jenis Pemeriksaan (Kategori) dan Bahagian Pemeriksaan kustom.
- **Dashboard Interaktif & Statistik Operasi**:
  - Reten jenis pemeriksaan bersepadu (*Consolidated Pivot Table*).
  - Kiraan pergerakan pesakit (Jalan Kaki, Wheelchair, Trolley, dan RME).
  - Perbandingan fasiliti sendiri vs luar serta pecahan mengikut setiap fasiliti rujukan individu.
  - Statistik Demografi (Jantina & Bangsa) dan inventori Consumables (CD & Filem X-Ray).
- **Matriks Reten PHRIS Bulanan**: Matriks reten 12 bulan (JAN - DIS) dengan kiraan automatik mengikut kategori rasmi KKM.
- **Keselamatan Data Tempatan (Offline-First)**: "Excel as a Database" kalis *blackout*, sokongan SQLite tempatan, giliran sinkronisasi (*offline sync queue*), dan arkib sandaran harian automatik.
- **Kompilasi Cross-Platform**: Pemasang sedia ada untuk macOS (`.dmg`) dan Windows (`.exe` / `.zip`).

---

## 2. Struktur Fail & Folder

```text
├── .github/
│   └── workflows/
│       └── build-release.yml       # GitHub Actions CI/CD (Auto-compile .dmg & .exe)
├── Daftar_Radiologi/
│   ├── DaftarRadiologi.py          # Enjin Aplikasi Utama (Flask & PyWebView)
│   ├── DaftarRadiologi.spec        # Konfigurasi PyInstaller Cross-Platform
│   ├── core/                       # Modul Teras (config, excel_engine, phris_engine, export_engine, db_engine, backup_engine)
│   ├── icon/                       # Ikon Aplikasi (.ico & .icns)
│   ├── routes/                     # Blueprint Flask (registration, patients, dashboard, phris, settings, export)
│   ├── templates/                  # Templat Antara Muka HTML
│   ├── static/                     # Fail CSS, JS (Chart.js), dan Aset Grafik
│   └── template.xlsx               # Templat Master Excel 12-Bulan (PER.SS-RA 101)
├── scripts/
│   ├── compile_and_setup_mac.sh    # Skrip kompilasi macOS tempatan
│   └── compile_and_setup_win.ps1   # Skrip kompilasi Windows tempatan
├── Install_Mac.command             # Skrip pelancaran macOS
├── Install_Windows.bat             # Skrip pelancaran Windows
├── requirements.txt                # Dependensi Python
└── README.md                       # Dokumentasi Projek
```

---

## 3. Panduan Pemasangan & Kompilasi

### A. Muat Turun Pemasang Terus (Recommended)
Pengguna boleh terus memuat turun pemasang rasmi dari bahagian [GitHub Releases](https://github.com/rdgr-szal/daftar-radiologi-v2/releases):
- 🍏 **macOS**: Muat turun `DaftarRadiologi-v2-macOS.dmg`, buka fail dan seret fail aplikasi ke folder *Applications*.
- 🪟 **Windows**: Muat turun `DaftarRadiologi-v2-Windows.zip`, ekstrak folder dan jalankan `DaftarRadiologi.exe`.

### B. Menjalankan Melalui Source Code (Developer Mode)
1. Bina Virtual Environment:
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate  # macOS/Linux
   # atau .venv\Scripts\Activate.ps1 bagi Windows
   ```
2. Pasang dependensi:
   ```bash
   pip install -r requirements.txt
   ```
3. Jalankan aplikasi:
   ```bash
   cd Daftar_Radiologi
   python3 DaftarRadiologi.py
   ```
4. Buka pelayar web pada `http://127.0.0.1:5005` (atau aplikasi akan dibuka terus melalui tetingkap desktop native PyWebView).