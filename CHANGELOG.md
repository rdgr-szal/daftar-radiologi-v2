# Changelog - Buku Daftar Radiologi V2

All notable changes to the **Buku Daftar Radiologi V2** project are documented in this file.

---

## [v2.1.5] - 2026-08-12

### 🚀 Major Feature & User Experience Updates

#### 1. Custom Starting X-Ray Sequence Number
- **Configurable Starting No.**: Added custom starting X-Ray sequence number configuration under **Settings > Modalities**. Facilities can now set custom sequence counters per modality room or start from a specific sequence index for the new year.

#### 2. Legacy Data Migration Tool
- **Historical Record Import**: Added legacy data import engine accessible via **Settings > System > Backup**. Allows importing legacy Excel records from older systems directly into the V2 database and PER.SS-RA 101 archives with duplicate checking.

#### 3. Interactive Onboarding Wizard
- **Guided Initial Setup**: Redesigned initial setup into an interactive 3-Step Wizard in full English. Guides new users step-by-step through facility profiling, staff roster setup, referral clinic configuration, and quick feature orientation.

#### 4. Smart Windows Setup Installer (Inno Setup)
- **Installer Package**: Built a native Windows installer (`Setup.exe`) via GitHub Actions CI/CD with automatic Desktop Shortcuts and Start Menu creation.
- **Smart Upgrade & Data Protection**: Automatically detects existing installations, offering *Upgrade*, *Repair*, or *Uninstall* options while strictly protecting patient records inside the `Pendaftaran/` directory from being overwritten or deleted.

#### 5. Minor UI & Visual Fixes
- **New Mascot & Brand Assets**: Updated application logo, icons, and favicons to the new bee mascot image.
- **Nav Branding Layout**: Enhanced sidebar navigation brand text styling to prevent "Daftar Radiologi" title clipping.
- **System Updates Badge**: Refined GitHub update check badge placement so update notices appear strictly under **Settings > System > Updates**.
- **Dark Clear Queue Action Button**: Styled DICOM Worklist "Clear Queue" button with dark red contrast styling for clearer action hierarchy.

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
