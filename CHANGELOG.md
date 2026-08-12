# Changelog - Buku Daftar Radiologi V2

All notable changes to the **Buku Daftar Radiologi V2** project are documented in this file.

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
