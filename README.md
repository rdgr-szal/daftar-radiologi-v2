# Sistem Input Buku Daftar Radiologi & Penjanaan Reten PHRIS (PER.SS-RA 101 Compliance)

Aplikasi berasaskan local web (localhost offline) yang direka khas untuk Unit Radiologi Klinik Kesihatan (KKM). Sistem ini membantu staf Radiologi mendaftarkan pesakit secara digital dan menjana reten statistik dalam format PHRIS secara automatik daripada pangkalan data fail Excel rasmi tanpa memerlukan sambungan internet.

---

## 1. Konsep Asas Aplikasi

Sistem ini dibina berasaskan konsep **"Excel as a Database"** bagi memenuhi syarat audit keselamatan KKM yang melarang penyimpanan data pesakit di dalam cloud.

```
[Borang Web (HTML/JS)] ➔ [Flask Backend (Python)] ➔ [openpyxl] ➔ [Fail Excel Tempatan (.xlsx)]
```

* **Penyimpanan Terus (Write-Through):** Tiada penggunaan cache memori (memory caching) atau pangkalan data kompleks (SQL). Setiap kali data pesakit dihantar, Python terus menulis baris baharu ke dalam helaian Excel harian (`1` hingga `31`) mengikut tarikh semasa, kemudian menyimpan (`save`) dan menutup fail tersebut serta-merta. Ini menjadikannya **kalis blackout**; jika bekalan elektrik terputus, data sebelum blackout dijamin selamat di cakera keras.
* **Auto-Increment Pintar:** Nombor X-ray seterusnya dikesan secara dinamik. Apabila borang dibuka, Python akan mengimbas helaian harian terkini dari bawah ke atas untuk mencari rekod terakhir dan melakukan operasi `Nombor Terakhir + 1`. Jika staf mengisi Excel secara manual semasa sistem ditutup, sistem akan mengesan baris manual tersebut dan menyambung reference dengan betul apabila dibuka semula.
* **Mesra Audit (Self-Contained Audit):** Sistem sync nama klinik asal dan rujukan terus ke dalam sheet `"Masterlist"`, `"Data Penuh"`, `"BIL PT"`, dan `"Reten Harian"` di dalam Excel. Jadi, jika pihak auditor membuka fail Excel secara manual tanpa menggunakan aplikasi ini, semua format, formula, dropdown, dan rumusan Excel tetap berfungsi dengan sempurna.

---

## 2. Struktur Fail & Folder

```text
├── Daftar_Radiologi/
│   ├── DaftarRadiologi.py          # Enjin Backend Flask & Logik Operasi Excel
│   ├── templates/                  # Template HTML (borang, dashboard, setup, data_penuh, reten)
│   ├── static/                     # Aset CSS/JS (termasuk Chart.js versi offline)
│   └── Pendaftaran/                # Folder Pangkalan Data Excel (diabaikan oleh .gitignore)
│       └── [Tahun]/
│           └── [No_Bulan]_[Bulan]/
│               └── [Tahun] [Bulan] PER.SS-RA 101 [SINGKATAN].xlsx
├── scripts/
│   ├── compile_and_setup_mac.sh    # Skrip kompilasi untuk macOS (PyInstaller)
│   └── compile_and_setup_win.ps1   # Skrip kompilasi untuk Windows (PyInstaller)
├── requirements.txt                # Senarai library luaran yang diperlukan
└── README.md                       # Dokumentasi Projek
```

---

## 3. Panduan Developer (Untuk Semakan / Review)

### Keperluan Pembangunan:
* **Python 3.10+**
* Library utama: `Flask` (Web Server), `openpyxl` (Manipulasi fail Excel), `pyinstaller` (compile as exe/app).

### Cara Menjalankan Mod Pembangunan (Dev Mode):
1. Bina Virtual Environment & Install dependacies:
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   pip install -r requirements.txt
   ```
2. Jalankan Aplikasi:
   ```bash
   cd Daftar_Radiologi
   python3 DaftarRadiologi.py
   ```
3. Buka browser: `http://localhost:5005`

### Logik Utama di Dalam [DaftarRadiologi.py](file:///Users/SZAL/Documents/Business/Web%20Dev/RADIOLOGY/Kesihatan%20Awam/OFFLINE%20BUKU%20DAFTAR/Daftar_Radiologi/DaftarRadiologi.py):
* `load_config()` & `save_config()`: Read/write local file `config.json` untuk tukar nama klinik dan senarai klinik rujukan secara dinamik tanpa ubah suai kod.
* `sync_excel_masterlists()`: Melakukan penyusunan semula fail Excel apabila nama/singkatan klinik diubah, termasuk menamakan semula nama fail di dalam disk dan menulis data konfigurasi ke dalam sel-sel Excel (`A13`, `A26`, `D2`, `H2` dll).
* `get_next_xray_no()`: Logik carian nombor X-ray terakhir berasaskan imbasan lajur C (Nombor X-ray) daily sheet terakhir.
* `/api/dashboard-data`: Agregasi dinamik data pesakit daripada fail Excel harian untuk diplotkan ke graf Chart.js di halaman utama dan halaman Analisis Rujukan.