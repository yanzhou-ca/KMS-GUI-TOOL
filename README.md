# 🚀 KMS Activation Tool — Release v1.0  
*A Modern, Secure, and Reliable GUI for KMS-Based Windows & Office Activation*  
**Released**: November 6, 2025  

---

## ✅ Overview  
The **KMS Activation Tool** is a self-contained, PowerShell-based graphical utility designed for IT professionals to quickly and securely activate **Windows and Office** products against a Key Management Service (KMS) host. Built with robustness, compatibility, and usability in mind, it eliminates command-line complexity while ensuring reliability across legacy and modern environments.

No installation required — just **run as Administrator**, enter your KMS server, and activate.

---

## 🔑 Key Features

### ✅ Universal Compatibility
- ✅ **OS Support**: Windows Vista → Windows 11 | Server 2008 → Server 2025  
- ✅ **Office Support**: Office 2010 → Office LTSC 2024  
- ✅ **PowerShell**: v2.0+ (Windows 7 SP1 and newer)  
- ✅ **Architecture**: Fully supports 32-bit and 64-bit systems (uses `Sysnative` for redirection safety)

### ✅ Self-Contained Hybrid Launcher
- 📦 Single `.bat` file — no installers, no dependencies  
- ⚡ Auto-bypasses PowerShell execution policy  
- 🔐 No temp files, no external scripts, no elevation prompts beyond initial UAC  
- 🚀 Portable — runs from USB, network share, or local drive

### ✅ Professional GUI Experience
- 🖥 Clean, responsive, dark-themed WPF interface  
- 📁 Tabbed layout: **Windows** | **Office**  
- ✅ Real-time validation — *Activate* button enabled only when inputs are valid  
- 📏 Fully resizable — log panel adjustable via `GridSplitter`  
- 🧵 Fully asynchronous — **no hangs or "Not Responding"**

### ✅ Smart & Secure Activation
- 🛠 Uses the **right tool for the job**:
  - `slmgr.vbs` for Windows  
  - `ospp.vbs` for Office  
- ✅ Correct command syntax (e.g., `/sethst:server`, not `/sethst server`)  
- 🚨 Intelligent error handling:
  - `0xC004F074` → KMS host unreachable  
  - `0xC004F038` → Insufficient activation requests  
- 🔒 **Full key masking** — only last 5 characters shown in logs  
- 🧹 Clean log output — no `---Processing---` or separator spam

### ✅ Comprehensive Product Coverage
- 🧾 **425+ official GVLKs** preloaded, including:
  - `Windows Server 2025`: Standard, Datacenter, Azure Edition  
  - `Windows 11`: Enterprise G, Pro for Workstations, Education N  
  - `Office LTSC 2024`: Professional Plus  
  - Legacy: Server 2008, Office 2010, Windows Vista  
- 📅 Products sorted **newest → oldest** for fast access  
- 🗂 Logical grouping by version and edition (Datacenter → Standard → Essentials)

---

## 🛠 Usage

1. **Right-click** `kms-gui-full.bat` → **Run as administrator**  
2. Enter KMS server address (e.g., `kms.yourhost.com:1688` or `192.168.0.1`)  
3. Switch to **Windows** or **Office** tab  
4. Expand a product category and select a specific edition  
5. Click **ACTIVATE**  
6. ✅ View real-time status in the log panel  

> 💡 **Pro Tip**: The *Activate* button auto-enables only when a valid KMS server and product are selected.

---

## 🔒 Security & Compliance
- 🔐 Requires admin rights (as mandated by Windows activation APIs)  
- 📜 Keys sourced from [Microsoft’s official GVLK documentation](https://learn.microsoft.com/en-us/windows-server/get-started/kmsclientkeys)  
- 🙅‍♂️ **No telemetry, no internet calls, no external dependencies**  
- 📝 Logs never expose full keys — compliant with security policies

---

## 📜 License & Compliance
> This tool is intended for use in environments with valid **Microsoft Volume Licensing** agreements and a properly configured, authorized KMS infrastructure.  
> ⚠️ Unauthorized use violates the Microsoft Software License Terms.

---

## 🙏 Acknowledgments
- Microsoft — for transparent GVLK documentation  
- PowerShell & WPF communities — for open, robust patterns  
- Internal IT teams — for real-world validation and feedback

---

**Prepared by**: Y.Z — IT Automation & Systems Engineering  
