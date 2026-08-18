# Changelog - Buku Daftar Radiologi V2

All notable changes to the **Buku Daftar Radiologi V2** project are documented in this file.

## [v2.1.6] - 2026-08-13

### 🚀 Bug Fix, CD Label Refactor & Platform Optimization Release

#### 1. Per-Patient CD Thermal Label Printing Refactor
- **Consolidated Patient Sticker**: Refactored CD thermal sticker label printing logic so 1 patient visit generates 1 consolidated CD label merging all X-Ray numbers (e.g. `NO. X-RAY: 0001 – 0005`) and all examination parts (e.g. `EXAM: CHEST PA + C-SPINE`).
- **Updated Label Hierarchy**: Rearranged label layout order to `TARIKH` → `ID` → `NAMA` (top section) and `NO. X-RAY` → `EXAM` (bottom section).
- **Cross-Platform Print Dialog Fix**: Replaced `window.open()` print window approach with an invisible iframe injection technique, ensuring instant, popup-blocker-free thermal printing on both macOS WKWebView and Windows WebView2.

#### 2. Patient List Date Picker & Event Handler Fixes
- **JavaScript Syntax Repair**: Fixed JavaScript parsing crash caused by escaped backticks inside template literal print code, restoring button clickability (`Week/Month/Year` tabs, `Print CD Label`, `Print / PDF`) across `patient_list.html`.
- **Native Date Picker Trigger**: Applied native `.showPicker()` popovers to date inputs (`#dayDateInput`, `#edit_lmp`, print modal inputs) with explicit cursor pointers.

#### 3. Windows System Freeze & Auto-Update Resolution
- **Temp Directory Traversal Guard**: Moved temporary file handling to OS temp directory (`tempfile.gettempdir()`) and updated `create_zip_backup()` to ignore `.zip` and `.tmp` files to prevent recursion loops and Windows freezing.
- **In-App Auto-Update Extract Fix**: Added root folder path cleaning for downloaded update zip archives and added startup patch handler `apply_pending_updates()` in `DaftarRadiologi.py`.

#### 4. Universal 5-Stage DICOM Connection Test Engine
- **Universal Modality Verification**: Implemented a comprehensive 5-stage diagnostic test (`test_dicom_connection`) in `core/dicom_engine.py` compatible with all modality console vendors (*Carestream, Konica Minolta, GE, Fuji, Agfa, Philips, Siemens*, etc.).
- **Multi-Layer Connectivity Inspection**:
  1. *ARP & LAN Cache*: Quick inspection to confirm physical LAN/gateway reachability.
  2. *Local Server Listener*: Verifies local DICOM MWL server is active on port `104` (or configured port).
  3. *TCP Network Connection*: Validates socket reachability to Console IP and Port.
  4. *DICOM C-ECHO Verification*: Association ping testing AE Title recognition (`0x0000`).
  5. *C-FIND MWL Inquiry*: Accurately distinguishes console SCU worklist receiver mode from SCP worklist provider mode.
- **Diagnostics UI & Dedicated Endpoint**: Added `POST /api/dicom/test-connection` endpoint and a structured diagnostic card in **Settings > Integration** detailing status checks with root-cause troubleshooting guidance.

#### 5. Excel Template & Standard 25-Column Format Realignment
- **25 Standard Columns Order**: Realigned Excel register output to match explicit sequence (`TARIKH`, `LMP`, `Bil Kes`, `Nombor X-ray`, ...).
- **KKM Styling**: Applied `#2F6EBA` (Blue) header fill to Rows 1 & 2 with white text, and `#FFFF00` (Yellow) fill to Row 4 headers with bold text.
- **Programmatic Generator & Backward Compatibility**: Added `generate_master_template_file()` for instant template generation with fallback support when reading legacy register files.

---

## [v2.1.5] - 2026-08-12

### 🚀 Maintenance & Feature Release

#### 1. Klinik Kesihatan (KK) Preset Modality Enforcer
- **General X-Ray Only**: Fixed setup wizard and settings modality preset for KK facilities to default strictly to `GENERAL RADIOGRAPHY`, unchecking Mobile X-Ray, Ultrasound, and Dental automatically.

#### 2. Drag & Drop File Upload Guard (PyWebView / Windows)
- **Prevent WebView Freeze**: Added global window drag/drop guards and dedicated dragover/drop event listeners to `#importDropZone` and `#updateDropZone` to prevent PyWebView / WebView2 page navigation or freezing when dropping files.

#### 3. Enlarged Native Application Icon
- **Full-Bleed Multi-Resolution ICO**: Trimmed empty transparent padding from `logo.png` and regenerated `daftarradiologi.ico` and `favicon.ico` containing sizes `16x16` up to `256x256` for native Windows Desktop, Taskbar, and Start Menu sizing.

#### 4. Thermal CD Sticker Label Printer Feature
- **CD Disc Surface Stickers**: Added interactive CD Label Printing tool in `patient_list.html` with preset roll sticker dimensions (50x30mm, 60x40mm, 50x50mm, circular CD ring), live interactive preview, font scaling, field toggles, and thermal printer CSS `@page` layout.

#### 5. Strict Emoji-to-SVG UI Refactoring
- **Clean Interface Standard**: Replaced all application emoji icons across HTML templates with lightweight, crisp SVG icons.

---

## [v2.1.4] - 2026-08-12

### 🚀 Major Feature & User Experience Updates

#### 1. True Single-Page AJAX On-The-Fly Filtering (Zero Page Refresh)
- **Instant Filter Switching**: Period switcher tabs (*Day*, *Week*, *Month*, *Year*), date inputs, and dropdowns on sticky header bars in `dashboard.html` and `patient_list.html` now update live via AJAX without browser page reloads.
- **Fixed Scroll Position**: Browsing long patient lists or census tables maintains the exact scroll position when switching dates or periods.

#### 2. Recalculated PHRIS Retention Statistics (PER.SS-RA 101 Compliance)
- **Patient Ethnicity (Demographics)**: Now accurately counts **Unique Patients** per month rather than examination procedure cases.
- **Arrival Categories**: Accurately counts **Unique Patients** arriving by *Trolley*, *Wheelchair*, *Rujuk Terus (Klinik Luar)*, and *Klinik/OPD (Walking / Walk-in)*.
- **General Examinations**: Accurately counts total **Examination Cases & Procedures Done** (including split bilateral studies).

#### 3. Standalone Windows Offline PC Deployment
- **Embedded Web Engine Fallback**: Bundled PySide6 QtWebEngine embedded Chromium browser runtime into the executable package for Windows PCs operating without internet access and without pre-installed Edge WebView2.
- **Auto-Embedded Bootstrapper**: Includes offline `MicrosoftEdgeWebview2Setup.exe` installer for automatic silent installation if WebView2 runtime is missing.

#### 4. UI & Template Refinements
- **Form ID Field**: Renamed `No. IC / Pasport` to **`ID`** on patient registration forms.
- **LMP Not Applicable Toggle**: Added `Not Applicable (N/A)` switch for female patient LMP input fields.
- **Settings & DICOM Echo**: Translated setting buttons to English (`+ Add`, `+ Add Type`, `+ Add Consumable`) and added feedback UI for DICOM Echo test connection.
- **Print Template**: Added fallback print trigger and `&larr; Kembali` navigation button on PDF/Print pages.

---

## [v2.1.3] - 2026-08-12
- Added DICOM MPPS SCP server support and Reject Analysis integration.
- Added dedicated DICOM MPPS Reject Logs tab in System Settings.
- Refined KK preset modality defaults to General Radiography.
