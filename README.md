# Radiology Registration & PHRIS Monthly Census System V2 (PER.SS-RA 101 Compliance)

<img width="1280" height="640" alt="image" src="https://github.com/user-attachments/assets/263d0179-ec07-4f49-ad25-07c753a0e8e2" />


Daftar Radiologi V2 is an offline desktop application designed for Radiology Units in Ministry of Health Malaysia (KKM) health clinics. The application assists radiographers in logging daily patient examinations, monitoring clinic statistics, and generating monthly PHRIS census reports directly from local Excel files, structured according to PER.SS-RA 101 standards.

---

## 1. Key System Features

- **Patient Registration Form**: Form fields aligned with the standard PER.SS-RA 101 register format, featuring automatic parsing of age, gender, and nationality from MyKad or Passport entries.
- **Consistent Date Formatting (`DD/MM/YYYY`)**: Standardised date display across registration forms, patient directories, filter pickers, and exported reports.
- **Examination Categories**: Pre-configured categories for common X-ray modalities and body sub-regions, with options to register custom examination descriptions as required by the facility.
- **Operational Summary Dashboard**:
  - Consolidated examination count tables.
  - Patient mobility summaries (Ambulatory, Wheelchair, Trolley, and RME).
  - Referral breakdown (Internal vs. External referrals, with counts per referring clinic).
  - Demographic distribution (Gender and Ethnicity) and basic consumable tracking (CDs and X-ray films).
- **Automated Monthly PHRIS Census**: 12-month statistical matrix (JAN – DEC) automatically calculated according to official KKM reporting guidelines.
- **Local File System Storage**: Direct read and write operations on local `.xlsx` files with daily automatic backups created locally.

---

## 💡 Technical Notes & Architecture Rationale

### Design Considerations:

1. **Local Data Storage & Offline Operation**:
   - Patient records are stored locally on the clinic workstation in standard Excel (`.xlsx`) files. This approach eliminates dependence on cloud infrastructure or active internet connections, allowing staff to access and backup spreadsheet records directly for administrative audits.
   - *Data Management*: To maintain consistency during heavy daily use, the application complements local Excel storage with a lightweight local index and automatic daily backup routines.

2. **File Saving Behaviour**:
   - Patient entries are saved directly to disk upon submission (`openpyxl`). Synchronous file writing helps protect recorded entries against loss in the event of unexpected power interruptions on clinic PCs.

3. **System Resource Efficiency**:
   - The application uses native desktop window rendering (PyWebView) rather than bundling a full web browser runtime. This maintains a small installation size and low RAM footprint, allowing smooth operation on routine clinic computers.

4. **GUI Launcher Fallbacks**:
   - To ensure stable operation across various Windows and macOS environments, the application includes a multi-tiered launcher fallback:
     1. **PyWebView** (Native OS window launcher).
     2. **PySide6 QtWebEngine** (Embedded GUI window fallback).
     3. **Local Web Browser Fallback** (Serves interface at `http://127.0.0.1:5005` in default browser).

---

## 2. Directory Structure

```text
├── .github/
│   └── workflows/
│       └── build-release.yml       # GitHub Actions CI/CD (Compiles macOS & Windows binaries)
├── Daftar_Radiologi/
│   ├── DaftarRadiologi.py          # Application desktop entry point (Flask + PyWebView / PySide6)
│   ├── DaftarRadiologi.spec        # PyInstaller packaging configuration
│   ├── core/                       # Core modules (config, excel_engine, phris_engine, export_engine, db_engine, backup_engine)
│   ├── icon/                       # Application icons (.ico & .icns)
│   ├── routes/                     # Flask Blueprints (registration, patients, dashboard, phris, settings, export)
│   ├── templates/                  # HTML interface templates
│   └── static/                     # Styling (CSS), JavaScript (Chart.js), and asset files
│   └── template.xlsx               # Standard 12-Month Excel Master Template (PER.SS-RA 101)
├── requirements.txt                # Python dependencies
└── README.md                       # Documentation
```

---

## 3. Installation & Setup Guide

### Option A: Direct Installer Download (Recommended for End Users)
Pre-built packages are available on [GitHub Releases](https://github.com/rdgr-szal/daftar-radiologi-v2/releases):

- 🍏 **macOS**:
  - **Apple Silicon (M1 / M2 / M3 / M4)**: Download `DaftarRadiologi-v2-mac-applesilicon.dmg` or `DaftarRadiologi-v2-mac-applesilicon.zip`.
  - **Intel Macs (macOS 10.15 Catalina or newer)**: Download `DaftarRadiologi-v2-mac-intel.dmg` or `DaftarRadiologi-v2-mac-intel.zip`.
- 🪟 **Windows**:
  - **Windows 10 / Windows 11**: Download `DaftarRadiologi-v2-win10.zip`, extract the ZIP file, and double-click `DaftarRadiologi.exe`.
  - **Legacy Windows (Windows 7 / 8 / 8.1)**: Download `DaftarRadiologi-v2-win-legacy.zip` (built with Python 3.8 for older Windows systems).

### Option B: Developer / Source Code Mode
1. Create and activate a Python Virtual Environment:
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate  # macOS / Linux
   # or .venv\Scripts\Activate.ps1 on Windows
   ```
2. Install required Python packages:
   ```bash
   pip install -r requirements.txt
   ```
3. Run the application:
   ```bash
   cd Daftar_Radiologi
   python3 DaftarRadiologi.py
   ```
4. Access the application window or open `http://127.0.0.1:5005` in your web browser.

---

## 4. First-Time Security Prompt Instructions

As application binaries are distributed without a paid commercial certificate, your operating system may display a security prompt upon opening the file for the first time.

### macOS (Gatekeeper Prompt):
If macOS indicates that the application is from an unidentified developer:
1. Right-click (or press `Control` and click) `Daftar Radiologi.app` inside your `/Applications` folder.
2. Select **Open**, then click **Open** again on the confirmation prompt. macOS will remember this selection for subsequent launches.
3. Alternatively, you may remove the quarantine attribute via Terminal:
   ```bash
   xattr -cr /Applications/DaftarRadiologi.app
   ```

### Windows (SmartScreen Prompt):
If Windows SmartScreen displays a protective prompt when starting `DaftarRadiologi.exe`:
1. Click **More info**.
2. Click **Run anyway** to proceed with launching the application.
