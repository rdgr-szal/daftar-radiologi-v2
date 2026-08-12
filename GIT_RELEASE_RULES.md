# Git Release & Update Workflow Rules

This document outlines the standard operating procedures and rules for committing code, pushing updates to Git, packaging builds for macOS and Windows, and updating project documentation for **Daftar Radiologi V2**.

---

## 📌 1. Language & Communication Standards

- **Malaysian English Standard**: All Git commit messages, Pull Request descriptions, release notes, `README.md`, `CHANGELOG.md`, `UPDATE.txt`, `INSTALLATION.txt`, and developer documentation MUST be written in **Malaysian English**.
- **Tone & Quality**: Use natural flow sentences, clear professional tone, correct grammar, and accurate technical terminology suited for Ministry of Health Malaysia (KKM) healthcare IT environments.

---

## 🚀 2. Git Update & Documentation Rules

### A. Updating `README.md`
- Every Git push that introduces new features or bug fixes **MUST** update `README.md`.
- **Location Rule**: Update **ONLY** the bottom section of `README.md` under **`5. Release History & Version Notes`** (or `Latest Release & Version Notes`).
- Do not modify historical section entries unless correcting factual errors.

### B. Versioning & Release Notes Strategy
- **New Version Releases (e.g. `v2.1.4` -> `v2.2.0`)**:
  1. Add a new version header under `CHANGELOG.md` and `README.md`.
  2. List major feature updates, UI enhancements, and system changes with clear bullet points.
  3. Tag the Git commit (`git tag -a v2.2.0 -m "Release v2.2.0"`) and push tags (`git push origin --tags`).
  
- **Same Version / Maintenance Updates**:
  1. If pushing updates under the same version tag or making minor tweaks without major feature additions, record the changelog entry with the concise phrase **"minor changes fix"** or **"minor fix"**.
  2. Example entry:
     ```markdown
     ### 🚀 Version 2.1.4 (Maintenance Patch)
     - **minor changes fix**: Resolved minor table alignment and button spacing tweaks.
     ```

---

## 📦 3. Platform Build Specifications

### 🪟 Windows (Offline Native App Architecture)
To ensure clinic PCs operating without internet access can run the application seamlessly as a native desktop window:

1. **Embedded WebView Kit & PySide6**:
   - Compiles with **PyWebView** for native Edge WebView2 rendering.
   - Bundles **PySide6 (QtWebEngine)** as an offline embedded Chromium fallback when WebView2 is absent.
   - Embeds `MicrosoftEdgeWebview2Setup.exe` bootstrapper in the package for silent background setup.
2. **OS Version Compatibility Coverage**:
   - **Modern Windows (Windows 10 / Windows 11)**: Compiled using Python 3.11 x64 target.
   - **Legacy Windows (Windows 7 / 8 / 8.1)**: Compiled using Python 3.10 x64 target to maintain compatibility with older OS kernel dependencies.

### 🍏 macOS (Native WebKit Architecture)
1. **Built-in WebKit Runtime**:
   - macOS includes native **WKWebView (WebKit)** as part of the operating system.
   - macOS builds do **NOT** bundle PySide6 or external WebKit libraries, keeping installer packages (`.dmg` & `.zip`) compact and lightweight.
2. **Target Architectures**:
   - **Apple Silicon (M1 / M2 / M3 / M4)**: Native ARM64 build.
   - **Intel Macs (macOS 10.15 Catalina+)**: Legacy x86_64 build.

---

## 🔄 4. Update Files & In-App Patching Rules

### A. Update Package Construction
- When publishing an update package (e.g., `DaftarRadiologi-v2.1.4-update.zip`):
  - **Include**: Executable binaries, `templates/`, `static/`, `core/`, `requirements.txt`, and root config files.
  - **EXCLUDE ALWAYS**: The **`Pendaftaran/`** folder.

### B. Data Protection Mandate
- **NEVER** include the `Pendaftaran/` directory in update ZIP packages or installation patches.
- The `Pendaftaran/` folder stores the clinic's local Excel database (`BUKU DAFTAR XRAY {Year}.xlsx`) and SQLite operational index. Including or overwriting this folder will cause data loss for end users.

### C. Manual & In-App Patching Flow
- Users can apply updates directly within the app via **Settings > System > Updates** by uploading the patch ZIP, or by manually extracting the patch ZIP into their existing application directory.

---

## 📖 5. Complete Windows Installation & Update Guide

### 📥 A. Fresh Installation (Windows 10 / 11 and Windows 7 / 8 / 8.1)

1. **Download Package**:
   - For Windows 10/11: Download `DaftarRadiologi-v2-win10.zip`.
   - For Windows 7/8/8.1: Download `DaftarRadiologi-v2-win-legacy.zip`.
2. **Extract Archive**:
   - Right-click the downloaded `.zip` file and select **Extract All...**.
   - Choose a convenient location (such as `C:\DaftarRadiologi` or Desktop).
3. **Launch Application**:
   - Open the extracted folder and double-click **`DaftarRadiologi.exe`**.
4. **First-Time Security Prompt**:
   - If Windows SmartScreen displays a blue warning ("Windows protected your PC"), click **More info** and then select **Run anyway**.
5. **Initial Setup**:
   - The application will open in a native window. On first run, complete the facility name and radiographer defaults in the Setup screen.

---

### 🔄 B. Updating Existing Installation (Manual Patching Steps)

1. **Close Application**:
   - Ensure **Daftar Radiologi** is completely closed.
2. **Download Update Patch**:
   - Download the latest update ZIP package from GitHub Releases or developer link.
3. **Extract & Overwrite Files**:
   - Open your existing `DaftarRadiologi` folder.
   - Extract the contents of the update ZIP into your existing folder.
   - When prompted by Windows, select **Replace the files in the destination** (Overwrite).
   
   > ⚠️ **CRITICAL DATA SAFETY NOTICE**:
   > **DO NOT** delete or overwrite your existing **`Pendaftaran`** folder. Your `Pendaftaran` folder contains all recorded patient data in the annual Excel register (`BUKU DAFTAR XRAY {Year}.xlsx`).
   
4. **Restart Application**:
   - Double-click `DaftarRadiologi.exe` to launch the updated version. Your patient database and system configurations will remain fully intact.
