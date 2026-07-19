# 📖 Panduan Pemasangan & Update / Installation & Update Guide
### Daftar Radiologi

---

## 🇲🇾 Bahasa Melayu (Mixed)

Panduan ringkas untuk install, run, dan update app Daftar Radiologi di PC klinik anda.

### 🚀 1. Keperluan Awal (Prerequisites)
* Perlu ada **Python (3.10 ke atas)**.
* Skrip installer (Langkah 2) akan tolong check & auto-install Python jika tiada.

### 🛠️ 2. Pemasangan Baru (Fresh Installation)
1. Copy folder projek ini ke dalam komputer.
2. **Windows:** Double-click fail **`Install_Windows.bat`**.
3. **macOS:** Double-click fail **`Install_Mac.command`**.
4. Skrip akan pasang dependencies dan create folder `DaftarRadiologi_App` serta shortcut **`Daftar Radiologi`** di Desktop.
5. Selesai! Klik shortcut di Desktop untuk run app.

### 🔄 3. Cara Update Aplikasi (Untuk Pengguna)
Jika ada update/reka bentuk baru dari developer:
1. Pastikan app ditutup.
2. Dapatkan fail zip update daripada developer.
3. Extract dan **copy/overwrite** fail baru masuk ke dalam folder `DaftarRadiologi_App` di Desktop anda.
4. ⚠️ **AMARAN:** Jangan padam atau replace folder **`Pendaftaran`** sedia ada kerana ia mengandungi semua rekod pesakit anda!

### ⚠️ 4. Jika Tersekat Masalah Keselamatan OS (Security Issue)
#### **Untuk macOS (Access Privileges / Gatekeeper):**
* **Ralat Privilege:** 
  1. Buka **Terminal**.
  2. Taip `chmod +x ` (pastikan ada space di hujung).
  3. Drag fail **`Install_Mac.command`** masuk ke Terminal, tekan **Enter**. Run semula fail tersebut.
* **Unidentified Developer:** Hold butang **Control** + klik fail -> pilih **Open** -> klik **Open** sekali lagi.

#### **Untuk Windows (SmartScreen / PowerShell Block):**
* **SmartScreen:** Klik kanan fail **`Install_Windows.bat`** -> **Properties** -> tick kotak **Unblock** -> klik **Apply/OK**.
* **PowerShell Block:** Buka **PowerShell** (Run as Admin) -> run arahan ini -> cuba double-click semula fail bat:
  ```powershell
  Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
  ```

---
---

## 🇬🇧 English

Brief guide to install, run, and update the Daftar Radiologi app on your clinic's PC.

### 🚀 1. Prerequisites
* **Python (3.10 or higher)** is required.
* The installer scripts (Step 2) will automatically check and try to install it if missing.

### 🛠️ 2. Fresh Installation
1. Copy the project folder onto the PC.
2. **Windows:** Double-click **`Install_Windows.bat`**.
3. **macOS:** Double-click **`Install_Mac.command`**.
4. The script sets up the app, creates the `DaftarRadiologi_App` folder, and places a **`Daftar Radiologi`** shortcut on your Desktop.
5. All set! Double-click the Desktop shortcut to run.

### 🔄 3. How to Update the App (For Users)
When receiving an update from the developer:
1. Close the application.
2. Download the update zip file.
3. Extract and **copy/overwrite** the new files into your existing `DaftarRadiologi_App` folder on the Desktop.
4. ⚠️ **WARNING:** Do NOT delete or overwrite the **`Pendaftaran`** folder, as it contains all your patient records!

### ⚠️ 4. OS Security Troubleshooting
#### **macOS (Access Privileges / Gatekeeper):**
* **Privilege Error:**
  1. Open **Terminal**.
  2. Type `chmod +x ` (with a trailing space).
  3. Drag and drop **`Install_Mac.command`** into Terminal, press **Enter**. Re-run the file.
* **Unidentified Developer:** Hold the **Control** key + click the file -> select **Open** -> click **Open** again.

#### **Windows (SmartScreen / PowerShell Block):**
* **SmartScreen:** Right-click **`Install_Windows.bat`** -> **Properties** -> check **Unblock** -> click **Apply/OK**.
* **PowerShell Block:** Open **PowerShell** (Run as Admin) -> execute this command -> run the bat file again:
  ```powershell
  Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
  ```
