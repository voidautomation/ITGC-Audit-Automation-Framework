# 📊 SAP ITGC Audit Automation Tool

> **Automated solution to perform IT General Controls (ITGC) audits in SAP using Excel VBA and SAP GUI scripting**

---

## 📘 Overview

The **SAP ITGC Audit Automation Tool** is an Excel VBA-based automation framework integrated with SAP GUI that helps IT Auditors perform GRC monthly audits with:

- ✅ One-click execution for predefined ITGC controls
- 📸 Screenshot capturing and Word documentation
- 📁 Audit log tracking
- 🧠 Dynamic GUI for system/control selection

🎯 **Goal**: Minimize audit execution time, reduce manual errors, and maintain compliance documentation automatically.

---

## 🚀 Features

- ✅ SAP GUI scripting automation (fully integrated with Excel)
- 📄 Capture SAP screen steps as a visual audit trail
- 💾 Save full Word (`.docx`) report with dynamic naming
- 🧪 18+ ITGC controls (e.g., Privilege checks, SAP_ALL access, IXOS access)
- 🧠 Dynamic step description mapping per control
- 🗂️ Directory creation and validation
- 👨‍💼 SSO and credentialed login support
- 📋 Integrated logging and error messaging

---

## 🧩 Technologies Used

| Component          | Details                                   |
|--------------------|--------------------------------------------|
| VBA (Excel Macros) | Backend logic and UI (UserForms)           |
| SAP GUI Scripting  | Executes SAP transactions/T-codes          |
| Word Automation    | Generates visually documented `.docx` file |
| Windows API        | Manages PrintScreen, window behavior       |

---

## 🖥️ User Interface Screens

### 🔐 Login Form

> Dropdowns for:
> - Select ITGC Control
> - Select SAP System  
> - Date Range Selectors  
> - User ID / Password  
> - LOGIN Button triggers automation 🟢

### 📋 Audit Dashboard

> Components:
> - `XML SETTING` to load SAP system config
> - `OPEN LOGS` to review execution history
> - `RUN AUDIT` to execute end-to-end snapshot + Excel export
> - Status: Control coverage matrix by platform (ERP, GRC, BW, HANA…)

---

## 📌 ITGC Controls Supported

| Ref         | Description                                                              |
|-------------|--------------------------------------------------------------------------|
| ITGC01      | High-level Privilege Access Audit                                        |
| ITGC02      | Client Opening Audit                                                     |
| ITGC06      | Developer Key / Transport Change Audit                                   |
| ITGC07      | SAP_ALL and SAP_NEW Access for A & S Users                               |
| ITGC08      | Table Maintenance Access                                                 |
| ITGC10      | Alter Security Configuration                                             |
| ITGC12      | Job Admin Functions Audit (`SM37`)                                       |
| ITGC13      | Configuration Change Audit (`RZ10`, `RZ11`)                              |
| ITGC17      | Usage of custom T-CODE: `START_REPORT`                                   |
| ITGCGUIXT   | GUIXT role assignment validation                                         |
| ITGCIXOS    | IXOS Admin Role Check (e.g., `JR_R3_IXOS_ADMIN_EBS`)                    |

> Full description sourcing is dynamically mapped from `Sheet4` in Excel.

---

## ⚙️ How It Works

```vb
ExecuteSAPScripts()
 ├─ AttachToSAP()
 ├─ CheckLoginStatus()
 ├─ ExecuteITGCXX()    ' Based on dropdown
 ├─ Takescreenshot()   ' Steps logged visually
 └─ EnsureFolderPathExistsAndSave()

## 📂 Folder Structure
plaintext

├── /Forms
│   └── UserLoginForm.frm
├── /Modules
│   ├── ExecuteSAPScripts.bas
│   └── UtilityFunctions.bas
├── Sheet3 [Input Controls: Control ID, Dates, User Info]
├── Sheet4 [Descriptions, Checkpoints per Control]
└── /Reports
    └── [Generated .docx audit files]
## 🔧 Requirements
✅ Excel (Macros enabled)
✅ SAP Logon Pad (saplgpad.exe)
✅ SAP GUI Scripting (enabled on server + GUI)
✅ Word (for report generation)
🔐 Valid SAP credentials (UserID/Password or SSO)
✅ Instructions to Run
Clone or download the .xlsm macro workbook.
Open in Excel (Run as Administrator).
Click RUN AUDIT → UserLoginForm opens.
Fill in:
Control ID
SAP System
Date Range
User ID / Password
Click LOGIN and relax!
🩺 Troubleshooting
Issue	Resolution
SAP GUI not found	Check if SAP GUI is installed correctly and scripting is enabled
Login keeps failing	Credentials incorrect or system combo mismatched
Screenshots not captured	Clipboard access denied, try local environment
Folder save fails	Nonexistent path or permission issue
No steps/document saved	Ensure Word is installed, and SAP form loads properly

## 🧑‍💻 Developer
👨 Abhinav Kumar
Developer · SAP Automation Specialist · Excel VBA Enthusiast
