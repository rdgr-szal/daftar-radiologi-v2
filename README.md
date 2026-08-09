# Radiology Register & PHRIS Monthly Census System V2 (PER.SS-RA 101 Compliance)

A standalone offline desktop application engineered specifically for Radiology Units in Health Clinics under the Ministry of Health Malaysia (KKM). The system enables radiographers to digitally log daily patient examinations, monitor operational statistics on an interactive dashboard, and automatically generate official monthly PHRIS statistical reports directly from a local Excel database without requiring internet connectivity.

---

## 1. Key System Features

- **Full PER.SS-RA 101 Compliance**: Official KKM registry format with automatic parsing of MyKad / Passport numbers (calculating Age, Gender, and Nationality).
- **Standardized Date Formatting (`DD/MM/YYYY`)**: Uniform date format across registration forms, patient directories, filter pickers, and print exports.
- **SMRP Taxonomy & Custom Orderables**: Official KKM examination catalog with the ability to dynamically add custom examination categories and anatomy sub-regions.
- **Interactive Operational Dashboard**:
  - Consolidated examination pivot tables (*Single Integrated Reten Table*).
  - Patient mobilization tracking (Ambulatory / Walking, Wheelchair, Trolley, and RME).
  - Facility breakdowns (Internal vs. External referrals, with granular counts per referring facility).
  - Demographic breakdowns (Gender & Ethnicity) and consumable inventory tracking (CDs & X-ray Films).
- **Automated Monthly PHRIS Census Matrix**: 12-month statistical grid (JAN – DEC) automatically calculated according to official KKM guidelines.
- **Offline-First Data Security**: Blackout-proof "Excel-as-a-Database" architecture, local SQLite indexing, offline synchronization queue, and automatic daily backup archiving.
- **Cross-Platform Packaging**: Automated CI/CD builds for macOS (`.dmg`) and Windows (`.exe` / `.zip`).

---

## 2. Directory Structure

```text
├── .github/
│   └── workflows/
│       └── build-release.yml       # GitHub Actions CI/CD (Auto-compiles .dmg & .exe)
├── Daftar_Radiologi/
│   ├── DaftarRadiologi.py          # Main Desktop Entrypoint (Flask + PyWebView + PySide6 Launcher)
│   ├── DaftarRadiologi.spec        # Cross-Platform PyInstaller Spec Configuration
│   ├── core/                       # Core Engines (config, excel_engine, phris_engine, export_engine, db_engine, backup_engine)
│   ├── icon/                       # High-Resolution Application Icons (.ico & .icns)
│   ├── routes/                     # Flask Blueprints (registration, patients, dashboard, phris, settings, export)
│   ├── templates/                  # HTML Interface Templates
│   └── static/                     # CSS (Charcoal Dark Theme), JS (Chart.js), and Graphic Assets
│   └── template.xlsx               # Official 12-Month Excel Master Template (PER.SS-RA 101)
├── requirements.txt                # Python Dependencies
└── README.md                       # Documentation
```

---

## 3. Installation & Usage Guide

### Option A: Direct Installer Download (Recommended for End Users)
Download pre-compiled native installers directly from [GitHub Releases](https://github.com/rdgr-szal/daftar-radiologi-v2/releases):
- 🍏 **macOS**: Download `DaftarRadiologi-v2-macOS.dmg`, double-click the DMG file, and drag `Daftar Radiologi` to your `/Applications` folder.
- 🪟 **Windows**: Download `DaftarRadiologi-v2-Windows.zip`, extract the archive, and double-click `DaftarRadiologi.exe`.

### Option B: Developer / Source Code Mode
1. Create a Python Virtual Environment:
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate  # macOS / Linux
   # or .venv\Scripts\Activate.ps1 on Windows
   ```
2. Install required dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Launch the application:
   ```bash
   cd Daftar_Radiologi
   python3 DaftarRadiologi.py
   ```
4. Access the web interface at `http://127.0.0.1:5005` or use the automatic desktop native window.

---

## 4. Installation Security Warning Bypass

Because this software is built for internal clinic deployment without commercial vendor certificate signatures:

- **macOS (Gatekeeper Warning)**:
  1. Right-click (or `Control` + click) `Daftar Radiologi.app` in `/Applications`.
  2. Click **Open**, then click **Open** again on the prompt. *(macOS permanently remembers this decision)*.
  3. Alternatively, run `xattr -cr /Applications/DaftarRadiologi.app` in Terminal.

- **Windows 10/11 (SmartScreen Warning)**:
  1. Click **"More info"** on the blue popup.
  2. Click **"Run anyway"**.