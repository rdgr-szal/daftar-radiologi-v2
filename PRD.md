# Product Requirements Document (PRD)
## Sistem Buku Daftar Radiologi & Reten PHRIS (PER.SS-RA 101 & SMRP 2.0 Compliance V2)

**Versi**: 2.2.0  
**Status**: Diluluskan / Universal Multi-Facility & MyGovUC 2.0 Cloud Backup  
**Penerbit**: Unit Radiologi Fasiliti Kesihatan KKM (Klinik Kesihatan, PKD & Hospital)  
**Persekitaran Utama**: Windows 10/11 Desktop (dengan sokongan tambahan macOS)

---

## 1. Pengenalan & Objektif Sistem

Sistem Buku Daftar Radiologi V2 ialah aplikasi desktop tempatan (*offline-first*) yang direka khas untuk Unit Radiologi fasiliti Kementerian Kesihatan Malaysia (KKM) di seluruh negara.

Aplikasi ini bersifat **Universal & Skalabel**:
1. **Mod Klinik Kesihatan (KK / KKIA / PKD)**: Menumpukan kepada General Radiography, Ultrasound asas, Dental/OPG dengan **Matriks Reten PHRIS 12 Bulan (PER.SS-RA 101 Compliance)** aktif.
2. **Mod Hospital / Institut KKM**: Menumpukan kepada katalog penuh **SMRP 2.0 Orderables** (General X-Ray, Mobile X-Ray, Fluoroscopy, Mammography, Ultrasound, CT Scan, MRI, Angiography, Interventional, BMD, C-Arm, IVU). Menyokong sama ada pendaftaran **Keseluruhan Jabatan** atau **Modaliti Khusus Sahaja** (cth: *Daftar CT-Scan Sahaja* dengan penamaan fail Excel khusus).

---

## 2. Seni Bina Dwi-Storan & Backup Automatik MyGovUC 2.0 Drive

```text
┌────────────────────────────────────────────────────────────────────────┐
│                        TAUX / NATIVE DESKTOP WINDOW                    │
│   (PyWebView: Microsoft WebView2 di Windows / Apple WKWebView di Mac) │
└───────────────────────────────────┬────────────────────────────────────┘
                                    │ Local IPC / HTTP
┌───────────────────────────────────▼────────────────────────────────────┐
│                    FLASK MODULAR BACKEND (Python 3.11+)                │
│ ├── core/config.py        : Pengurusan Konfigurasi JSON & Path         │
│ ├── core/excel_engine.py  : Enjin openpyxl 1 Sheet = 1 Bulan           │
│ ├── core/db_engine.py     : Dwi-Storan & Sync Queue (DB + Excel)       │
│ ├── core/backup_engine.py : Auto-Backup MyGovUC 2.0 Drive & ZIP Engine│
│ ├── core/phris_engine.py  : Pengira Matriks Reten & PDF/Print Renderer  │
│ └── routes/               : Modular Flask Blueprints                   │
└─────────┬─────────────────────────┬──────────────────────────┬─────────┘
          │ Direct File I/O         │ Offline-First Sync       │ Auto / 1-Click ZIP
┌─────────▼─────────────┐ ┌─────────▼─────────────────┐ ┌──────▼──────────────────┐
│ STORAN PRIMER (EXCEL) │ │ STORAN SEKUNDER (DB)      │ │ MYGOVUC 2.0 DRIVE       │
│ - Fail .xlsx 12 Bulan │ │ - MyGovCloud@PDSA         │ │ - Google Workspace KKM  │
│ - 100% Kalis Blackout │ │ - On-Premise MySQL/Postgres│ │ - Auto Backup Harian    │
│ - Sedia Audit Fizikal │ │ - SQLite Intranet         │ │ - 1-Klik Restore ZIP    │
└───────────────────────┘ └───────────────────────────┘ └─────────────────────────┘
```

### 🔒 Pematuhan Matrik Keselamatan Sektor Awam:
Berdasarkan Pekeliling Keselamatan Data Sektor Awam:
- Rekod Pesakit & Data PII adalah **Data Terkawal**.
- Penggunaan **Public Cloud & Off-shore (luar negara)** adalah tidak dibenarkan bagi data rasmi terhad/sulit.
- Model yang dibenarkan: **MyGovCloud@PDSA, On-Premise Kerajaan, Local Intranet SQLite/PostgreSQL/MySQL**, dan **Akaun Rasmi MyGovUC 2.0 Google Drive (@moh.gov.my / @kkm.gov.my)**.

---

## 3. Integrasi Backup Automatik & Migrasi PC (MyGovUC 2.0 Drive)

### 3.1 Ciri-ciri Backup Automatik:
1. **Penjadualan Harian Automatik (`auto_daily_backup_check`):**
   Aplikasi secara automatik membungkus folder `Pendaftaran/` (semua fail Excel `.xlsx`, `config.json`, `radiologi_local.db`, `sync_queue.json`) ke dalam fail `.zip` pada pelancaran atau jam 5:00 petang setiap hari.
2. **Penyelarasan ke MyGovUC 2.0 Google Drive:**
   Fail `.zip` disalin secara terus ke folder Google Drive Desktop rasmi unit Radiologi (contoh: `G:\My Drive\Backup_Radiologi`).
3. **Pembersihan Automatik (Retention Policy):**
   Fail backup tempatan melebihi 30 hari dibersihkan secara automatik bagi menjimatkan storan PC.

### 3.2 Migrasi PC & Pemulihan Data (*1-Click Restore*):
- Jika PC Kaunter rosak, petugas hanya perlu memasang aplikasi pada PC baharu dan menekan butang **"🔄 Pulihkan Data dari Fail Backup (.zip)"** di menu Settings.
- Sistem akan mengabstrak dan memulihkan semula keseluruhan pendaftaran Excel & DB dalam beberapa saat.

---

## 4. Pematuhan Kriteria Audit KKM bagi Fungsi Edit & Batal

Berdasarkan **Garis Panduan Pendaftaran Pemeriksaan Radiologi KKM 2021**:
1. **Nombor X-Ray Kekal Bersiri (Tiada Lompang Nombor):**
   Nombor siri pendaftaran tahunan tidak boleh dipadamkan sehingga mewujudkan kelompangan (*gap*) dalam rekod audit.
2. **Prosedur Pembatalan (*Cancellation / Batal*):**
   - Di Excel: Kolom `COMMENT` dikemaskini dengan tag `[BATAL: <Sebab>] OLEH <Staff>`, dan baris diwarnakan penanda pudar untuk audit visual.
   - Di DB: Status dikemaskini kepada `BATAL` bersama sebab pembatalan dan nama staf pengesah.
   - Statistik Reten PHRIS & SMRP mengecualikan kes batal daripada kiraan beban kerja secara pintar.
3. **Prosedur Kemaskini (*Edit*):**
   - Pembetulan nama, IC, umur, klinik rujukan, atau operator direfleksikan secara serentak ke dalam Excel dan Database.

---

## 5. Ringkasan Pengesahan & Pengujian

- [x] Penyingkiran *footprint* statik KKBBS & pengaktifan `is_configured = False` secara lalai.
- [x] Wizard Persediaan Universal dengan sokongan Klinik Kesihatan & Hospital.
- [x] Dwi-Storan: Fail Excel tempatan + Pangkalan Data Sekunder (SQLite / Postgres / MySQL / Cloudflare D1).
- [x] Sistem Giliran Luar Talian (*Offline Sync Queue*) dengan pemulihan automatik.
- [x] Backup Automatik Harian & Penyepaduan MyGovUC 2.0 Google Drive.
- [x] Pemulihan Data 1-Klik (*One-Click Restore*) daripada fail `.zip`.
- [x] Pematuhan audit PER.SS-RA 101 bagi fungsi Kemaskini (Edit) dan Pembatalan (Cancel).
- [x] Penamaan fail Excel dinamik mengikut modaliti khusus hospital.
