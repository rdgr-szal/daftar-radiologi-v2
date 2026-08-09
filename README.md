# Radiology Register & PHRIS Monthly Census System V2 (PER.SS-RA 101 Compliance)

A standalone offline desktop application engineered specifically for Radiology Units in Health Clinics under the Ministry of Health Malaysia (KKM). The system enables radiographers to digitally log daily patient examinations, monitor operational statistics on an interactive dashboard, and automatically generate official monthly PHRIS statistical reports directly from a local Excel database without requiring internet connectivity.

---

## 1. Key System Features

- **Fast Data Entry**: Official KKM registry format with automatic parsing of MyKad / Passport numbers (calculating Age, Gender, and Nationality).
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

## 💡 Footnote & Technology Stack Rationale

### Why This Architecture Was Chosen:

1. **Strict KKM Data Governance & Compliance (Offline-First)**:
   - Ministry of Health Malaysia (KKM) policies strictly prohibit storing unencrypted patient Personally Identifiable Information (PII) on public cloud servers.
   - **"Excel-as-a-Database" Pragmatism**: Directly reading and writing to local `.xlsx` files guarantees 100% offline data sovereignty, zero cloud telemetry, and direct compatibility with physical clinical audit standards without requiring complex database installation on clinic workstations.
   - *Developer's Note on Data Architecture*: While I personally advocate for relational databases over Excel for core storage, this design pragmatically satisfies existing KKM audit workflows. To ensure solid data integrity, I implemented a hybrid resilience layer—incorporating a local SQLite index, LAN network database sync queue, and automated daily backup routines for extra data protection.

2. **Blackout-Proof Write-Through Reliability**:
   - Every patient registration executes an immediate write-and-close operation (`openpyxl`) to disk. If a sudden power outage occurs in a rural health clinic, zero data is lost.

3. **Ultra-Lightweight Desktop Footprint (PyWebView)**:
   - By utilizing native OS rendering engines (WebKit on macOS, WebView2 on Windows) via **PyWebView** instead of bundling a heavy Chromium runtime, the installer size remains **~35MB** and RAM usage stays **under 60MB**—enabling smooth operation on low-spec government desktop PCs.

4. **Multi-Tier Window Fallback Resilience**:
   - To guarantee zero application crashes on legacy or unpatched clinic PCs lacking pre-installed WebKit/WebView2 runtimes, the app features a 3-tier launcher:
     1. **PyWebView** (Native OS Window - Ultra-fast & 35MB).
     2. **PySide6 QtWebEngine** (Embedded Chromium Window).
     3. **Local Web Browser Fallback** (Serves `http://127.0.0.1:5005` in default browser).

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