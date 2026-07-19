# PROJEK: Sistem Input Buku Daftar Radiologi & Penjanaan Reten PHRIS Offline

## 1. Objektif & Kekangan Projek
* **Tujuan:** Membina sistem borang input (Form Layout) dan paparan data (Dashboard & PHRIS Preview) berasaskan web tempatan (offline) untuk unit Radiologi Klinik Kesihatan.
* **Kekangan Keselamatan (Audit KKM):** Sistem MESTI berjalan 100% offline (tiada Cloud/Google Sheets/Drive/MyGovUC 3.0) untuk menjaga kerahsiaan data pesakit (PII).
* **Aliran Kerja Pembangunan:** Dibangunkan dan diuji di persekitaran **Mac (macOS)**, kemudian dipindahkan ke **Windows PC (Klinik)** untuk dikompilasi menjadi satu fail `.exe` guna PyInstaller.

## 2. Seni Bina Sistem (Tech Stack)
* **Frontend:** HTML5, CSS3, JavaScript (Menggunakan Chart.js versi offline/local untuk graf).
* **Backend:** Python (Flask) berjalan sebagai Localhost Server.
* **Pangkalan Data:** Fail Microsoft Excel (`.xlsx`) sedia ada yang diwartakan (disimpan mengikut folder tahun/bulan). Python menggunakan library `openpyxl` atau `pandas` untuk membaca/menulis data secara senyap di belakang tabir.
* **Platform Pembangunan:** macOS
* **Persekitaran Ujian:** Flask local server (`http://localhost:5000` ganti port lain sekiranya port sudah digunakan oleh server lain).
* **Platform Produksi:** Windows PC (Fail binary `.exe`) User hanya double klik dan ianya akan terbuka automatically untuk terus register pesakit.

## 3. Struktur Fail & Folder Sasaran
```text
Daftar_Radiologi/
├── DaftarRadiologi.py          # Enjin Python Flask
├── templates/
│   ├── borang.html             # Borang input data harian pesakit
│   ├── dashboard.html          # Graf visual analisis beban kerja (Chart.js)
│   └── reten_preview.html      # Halaman PREVIEW KHAS untuk reten PHRIS
├── static/
│   └── js/
│       └── chart.js            # Library grafik (offline version)
└── Pendaftaran/
    └── [Tahun]/
        └── [No_Bulan]_[Nama_Bulan]/
            └── [Tahun] [Bulan] PER.SS-RA 101 KKRP.xlsx


4. Reka Bentuk Data & Aliran Proses (Data Flow)
Input Data Harian (borang.html):

Staf isi borang ➔ Python kesan tarikh semasa ➔ Buka fail Excel bulan berkenaan ➔ Tulis ke Sheet nombor hari bulan tersebut ('1', '2', ..., '31').

Auto-Increment Nombor X-ray: Python mesti semak rekod terakhir di baris/sheet sebelum ini secara automatik, buat operasi +1, dan paparkan sedia ada pada borang HTML tanpa input manual.

Halaman Preview Reten PHRIS (reten_preview.html):

Halaman ini memaparkan satu jadual matriks (Kategori vs Bulan: JAN - DIS) yang menjumlahkan data dari bulan 1 hingga bulan semasa untuk tujuan mudahkan staf key-in ke sistem PHRIS KKM.

Struktur Kategori PHRIS yang Perlu Dijana Dinamik oleh Python:

Bangsa Pesakit: Melayu, Cina, India, Bumiputera, Warga Asing, Lain-lain, Jumlah.

Kedatangan: Trolley, Wheelchair, Rujuk Terus, Bil. Pesakit (Klinik/OPD), Jumlah.

Pemeriksaan AM: Dada, Abdomen, Extremiti, Rangka Kepala, Spina Vertebra, Pelvis, Skeletal Survey, Dexa, OPG, Lain-Lain, Jumlah.

Pemeriksaan Lain: RME, PTB, Kes-Kes Lain, COVID, DM, Jumlah.

Temujanji: USG, Mammo, Lain-Lain, CT-Scan, Jumlah.

Consumable: CD-R, Filem 10X12, Filem 14X17, Jumlah.

Penolakan Imej: Over Exposure, Under Exposure, Double Exposure, Wrong Technique, Wrong Patient/Exam, No Primary/Wrong Marker, Collimation Error, Patient Movement, Patient Related Artifact, Equipment Fault, Detector/Imaging Plate, Image Artifact, Processing Fault, Miscellaneous, Jumlah Imej, Pengulangan, Peratusan (%).

5. Arahan Seterusnya
Bantu saya mulakan kod fasa pertama (Blueprint):

Bina fail DaftarRadiologi.py berasaskan Flask yang mempunyai fungsi route untuk Borang (/), Dashboard (/dashboard), dan Reten PHRIS (/phris).

Sediakan fungsi logik Python untuk membaca data mentah dari helaian-helaian kerja bulanan Excel dan menyusunnya (aggregate) menjadi format jadual PHRIS (JAN-DIS) seperti struktur di atas.

Sediakan struktur template reten_preview.html.