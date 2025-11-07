🚀 KMS Activation Tool – Release v1.0  
A Modern, Secure, and Reliable GUI for KMS-Based Windows & Office Activation  
Released: November 7, 2025  

✅ Overview  
The KMS Activation Tool is a self-contained, PowerShell-based graphical utility designed for IT professionals to quickly and securely activate Windows and Office products against a Key Management Service (KMS) host. Built with robustness, compatibility, and usability in mind, it eliminates command-line complexity while ensuring reliability across legacy and modern environments.  
  
No installation required — just run as Administrator, enter your KMS server, and activate.  

🔑 Key Features  
✅ Universal Compatibility  
Supports Windows Vista through Windows 11  
Supports Windows Server 2008 through 2025  
Supports Office 2010 through Office LTSC 2024  
Works on PowerShell 2.0+ (Windows 7 SP1 and newer)  
Fully compatible with 32-bit and 64-bit systems (uses Sysnative to avoid file redirection)  
  
✅ Self-Contained Hybrid Launcher  
Single .bat file — no dependencies, no installers  
Bypasses PowerShell execution policy automatically  
Launches GUI safely — no temp files, no external scripts  
Fully portable — run from USB, network share, or local disk  
  
✅ Professional GUI Experience  
Clean, responsive, dark-themed WPF interface  
Tabbed layout: Windows and Office products clearly separated  
Real-time validation — “Activate” button enabled only when inputs are valid  
Resizable window with adjustable log panel  
Full async operation — no “Not Responding” hangs  
  
✅ Smart & Secure Activation  
Uses correct tools per product:  
slmgr.vbs for Windows  
ospp.vbs for Office  
Correct syntax handling (e.g., /sethst:server, not /sethst server)  
Full error detection (e.g., 0xC004F074 = KMS unreachable)  
Never logs full product keys — only last 5 characters for security  
Clean, actionable log output — no ---Processing--- noise  
  
✅ Comprehensive Product Coverage  
425+ GVLKs preloaded — including:  
Windows Server 2025 (Standard, Datacenter, Azure Edition)  
Windows 11 Enterprise G / Pro for Workstations  
Office LTSC 2024 Professional Plus  
Legacy support: Windows Server 2008, Office 2010, Windows Vista  
Products sorted newest first for faster access  
Logical grouping by version and edition  
  
🛠 Usage  
Right-click KMS_Activation_Tool.bat → Run as administrator  
Enter your KMS server (e.g., kms.contoso.com or 192.168.2.17)  
Select a product from the Windows or Office tab  
Click ACTIVATE  
  
✅ Done — clear success/failure feedback in the log panel  
💡 Tip: The tool automatically validates inputs and enables the button only when ready.   
  
🔒 Security & Best Practices  
Requires administrator privileges (as needed for slmgr.vbs)  
Keys are embedded in script — no external downloads  
Full key masking in logs (compliant with security policies)  
No telemetry, no internet access, no external dependencies  

📜 License
This tool is intended for use with legitimate Volume Licensing agreements and a properly configured KMS infrastructure.  
⚠️ Misuse violates Microsoft Software License Terms.  

🙏 Acknowledgments  
Microsoft for official GVLK documentation  
PowerShell and WPF communities for robust open patterns  
Internal IT teams for real-world testing and feedback  
