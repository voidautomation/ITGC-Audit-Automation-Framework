📊 SAP ITGC Audit Automation Tool
SAP Audit Logo

Automated solution to perform IT General Controls (ITGC) audits in SAP using Excel VBA and SAP GUI scripting

📘 Overview
The SAP ITGC Audit Automation Tool is an Excel VBA-based automation framework integrated with SAP GUI that helps IT Auditors perform GRC monthly audits with:

One-click execution for predefined ITGC controls
Screenshot capturing
Report generation in MS Word
Audit log tracking
Dynamic GUI for system/control selection
🎯 Goal: Minimize audit execution time, reduce manual errors, and maintain compliance documentation automatically.

🚀 Features
✅ SAP GUI scripting automation (fully integrated with Excel)
📄 Capture SAP screen steps as visual audit trail
📁 Save full Word (.docx) report with dynamic naming
🧾 18+ ITGC controls (e.g., Privilege checks, SAP_ALL access, IXOS access)
🧠 Dynamic step description mapping per control
🗃️ Directory creation validation
🧑‍💼 SSO and credentialed login support
✅ Integrated logging and error messaging
🧩 Technologies Used
Technology	Description
VBA (Excel Macros)	Backend logic and UI forms
SAP GUI Scripting	Executes T-codes and automates forms
Word Automation	Report generation (via Word Object)
Windows API	Performs PrintScreen, window management, etc.
🎛️ User Interface Screens
🔐 Login Form:
Login UI

Select Control
Choose SAP System
From/To Dates
User ID / Password input
📋 Dashboard:
Audit Dashboard

XML Import
Log Viewer
RUN AUDIT button
Control definitions + matrix by system
🧪 ITGC Controls Supported
Code	Definition
ITGC01	High-level Privilege Access Audit
ITGC02	Client Opening Audit
ITGC06	Developer Key/Transport Change Audit
ITGC07	SAP_ALL and SAP_NEW user privileges
ITGC08	Table maintenance access
ITGC10	Alter Security Configuration
ITGC12	Job Administration Audit (SM37)
ITGC13	Change configuration RZ10/RZ11
ITGC17	Usage of T-CODE START_REPORT
ITGCGUIXT	GUIXT Role Check
ITGCIXOS	IXOS Admin Role Check
...	... and more
Configurations are mapped via Excel sheets and fetched during execution dynamically.

🚦 How It Works
1. Setup
Enable Excel Macros
Confirm SAP GUI Scripting is enabled
SAP GUI > Options > Accessibility & Scripting > Enable scripting
Ensure valid SAP credentials and Excel inputs
2. Execution Flow
vba

ExecuteSAPScripts()
 |- AttachToSAP()
 |- CheckLoginStatus()
 |- ExecuteITGCXX()
 |- Takescreenshot()
 |- Save Report (Word)
3. Output
Word file:
ITGC01 ERP Audit Report.docx

Screenshot logs and user remarks embedded per step

📁 File Structure
Bash

├── /Forms
│   └── UserLoginForm.frm
├── /Modules
│   ├── ExecuteSAPScripts.bas
│   ├── UtilityFunctions.bas
├── Sheet3 [Inputs: Control ID, Dates, Paths]
├── Sheet4 [Descriptions, Checkpoints]
└── /Reports
    └── [Auto-saved Word documents]
🛠️ Admin/Developer Notes
Setup Script Logging:
Uses LogMessage subroutine to save details locally
UploadLogs.UploadFileToFirebase (optional Firebase support)
Add New Control:
Write new Subroutine:
vba

Sub ExecuteITGCXX()
  ' Add SAP logic and screenshots
End Sub
Update ExecuteSAPScripts case selector
Add descriptions to Sheet4
🧯 Troubleshooting
Issue	Fix
SAP GUI not found	Ensure saplgpad.exe is installed and scripting enabled
No screenshots captured	Check PrintScreen key availability & local clipboard access
Login stuck	Validate credentials & SAP landscape configuration
Folder save error	Use full path or check folder permissions
🧑‍💻 Developer
👨‍💻 Developer: Abhinav Kumar
📧 Email: Kumar.abhinav227@gmail.com


📄 License
This project is for internal business use only. Redistribution or reuse for commercial purposes requires explicit authorization.
