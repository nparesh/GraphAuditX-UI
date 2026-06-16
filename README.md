# GraphAuditX UI

> **GraphAuditX Pro** — A WPF-based PowerShell tool for querying Microsoft 365 Purview Unified Audit Logs via Microsoft Graph API, O365 Management API, and Search-UnifiedAuditLog. Includes CSV analysis, HAR file parsing, query generation, and full GUI.

**Author:** Paresh Nhathalal — Microsoft Purview  
**Version:** 6.0 | June 2026

---

## 🚀 Quick Start (5 minutes)

```powershell
# 1. Clone the repo
git clone https://github.com/nparesh/GraphAuditX-UI.git C:\GraphAuditX

# 2. Run it
powershell.exe -ExecutionPolicy Bypass -File "C:\GraphAuditX\GraphAuditX-WPF.ps1"
```

---

## 📋 Prerequisites

| Requirement | Details |
|---|---|
| **OS** | Windows 10 / 11 / Server 2016+ |
| **PowerShell** | Windows PowerShell 5.1 (ships with Windows) — **NOT** PowerShell 7/pwsh |
| **.NET Framework** | 4.7.2+ (for WPF / PresentationFramework) |
| **Microsoft.Graph module** | Optional — only needed for live Graph API queries |
| **Entra App Registration** | Optional — only needed for live audit log queries |

---

## 📥 Step 1: Download

### Option A: Git Clone (recommended)

```powershell
git clone https://github.com/nparesh/GraphAuditX-UI.git C:\GraphAuditX
```

### Option B: Download ZIP

1. Go to https://github.com/nparesh/GraphAuditX-UI
2. Click **Code** → **Download ZIP**
3. Extract to `C:\GraphAuditX\`

---

## 📂 Step 2: Understand the Folder Structure

```
GraphAuditX-UI\
├── GraphAuditX-WPF.ps1              ← 🎯 Main WPF GUI application (launch this)
├── AuditOperationsGateConfig.csv    ← Workload/Operations config data (144 workloads)
├── GraphAuditX.psd1                 ← Module manifest
├── GraphAuditX.psm1                 ← Module loader
├── RecordTypeResolver.cs            ← Record type enum resolver (200+ types)
├── SETUP-GUIDE.md                   ← Detailed setup guide
├── GraphAuditX\                     ← Core PowerShell module
│   ├── GraphAuditX.psd1
│   ├── GraphAuditX.psm1
│   ├── Public\                      ← Exported functions
│   │   ├── Connect-GraphAuditX.ps1
│   │   ├── GraphAuditX.ps1         ← Main query function
│   │   └── Get-GraphAuditXStatus.ps1
│   └── Private\                     ← Internal helpers
│       ├── Connect-GraphAuditX.ps1
│       ├── Get-GraphAuditXToken.ps1
│       └── Invoke-GraphAuditXRequest.ps1
├── M365Api\                         ← O365 Management Activity API support
│   └── M365-ManagementAPI.ps1
└── Assets\                          ← GUI icons/images
    ├── logo.png
    └── watermark.jpg
```

---

## ⚙️ Step 3: Update Paths

Open `GraphAuditX-WPF.ps1` and update the **3 paths** near the top to match your install location:

```powershell
# Line ~3 - M365 Management API helper
. "C:\GraphAuditX\M365Api\M365-ManagementAPI.ps1"

# Line ~11 - Module import
Import-Module "C:\GraphAuditX\GraphAuditX\GraphAuditX.psd1" -Force

# Line ~12 - Config CSV
$configPath = "C:\GraphAuditX\AuditOperationsGateConfig.csv"
```

> **If you cloned to `C:\GraphAuditX`** these should already work. Otherwise replace with your actual path.

---

## 🔓 Step 4: Set Execution Policy

```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope CurrentUser
```

---

## ▶️ Step 5: Launch the Tool

```powershell
# Option 1: Direct launch
powershell.exe -ExecutionPolicy Bypass -File "C:\GraphAuditX\GraphAuditX-WPF.ps1"

# Option 2: From within PowerShell
& "C:\GraphAuditX\GraphAuditX-WPF.ps1"
```

The **GraphAuditX Pro** WPF window will appear.

---

## 🔑 Step 6: (Optional) Configure Entra App for Live Queries

To run **live audit log queries** against a tenant, you need an Entra App Registration:

### 6a. Create the App Registration

1. Go to [Azure Portal → Entra ID → App registrations](https://portal.azure.com/#view/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/~/RegisteredApps)
2. Click **New registration**
   - **Name:** `GraphAuditX`
   - **Supported account types:** Single tenant (or Multi-tenant for cross-tenant)
   - **Redirect URI:** Leave blank
3. Note your **Tenant ID** and **Application (Client) ID**

### 6b. Add API Permissions

Navigate to **API permissions** → **Add a permission**:

| API | Permission | Type |
|---|---|---|
| Microsoft Graph | `AuditLogsQuery.Read.All` | Application |
| Microsoft Graph | `AuditLog.Read.All` | Application |
| Office 365 Management APIs | `ActivityFeed.Read` | Application |
| Office 365 Management APIs | `ActivityFeed.ReadDlp` | Application |

Then click **✅ Grant admin consent for [your org]**

### 6c. Create Client Secret

1. Go to **Certificates & secrets** → **New client secret**
2. Set expiry (recommended: 12 months)
3. **Copy the secret value immediately** (it won't show again)

### 6d. Enter Credentials in Tool

In the GraphAuditX GUI:
- **Tenant ID:** paste your tenant GUID
- **Client ID:** paste the application ID
- **Client Secret:** paste the secret value

---

## 🎯 Features

### 🔍 Live Audit Query (Graph API)
- Select workloads, record types, operations, date range, users
- Queries via `Microsoft Graph API (beta/security/auditLog/queries)`
- Results in filterable/sortable DataGrid
- Supports all 144 workloads and 6,273 operations

### 📂 Upload & Analyze CSV
- Upload any Purview Audit export CSV
- **Analyze** generates report: top operations, workloads, record types, user activity
- **Query Generation** — generates ready-to-run scripts:
  - MS Graph Query (.ps1)
  - O365 Management API Query (.ps1)
  - PowerShell UAL Query (.ps1)

### 🔍 Upload & Analyze HAR
- Upload browser HAR file (network trace from Purview portal)
- Extracts all `AuditLogSearch` API calls
- Shows Query IDs, search parameters, date ranges, operations, users
- Reproduce any search someone performed in the portal

### 📋 Export
- Export results to CSV
- Open directly in Excel

---

## 💻 CLI Usage (Without GUI)

```powershell
# Import the module
Import-Module "C:\GraphAuditX\GraphAuditX\GraphAuditX.psd1" -Force

# Connect
$secret = Read-Host "Client Secret" -AsSecureString
Connect-GraphAuditX -TenantId "YOUR-TENANT-ID" -ClientId "YOUR-CLIENT-ID" -ClientSecret $secret

# Run a query (returns QueryId)
$queryId = GraphAuditX -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date) `
    -RecordTypeFilters @("ExchangeAdmin","SharePointFileOperation") `
    -OperationFilters @("FileAccessed","MailboxLogin")

# Check status
Get-GraphAuditXStatus -QueryId $queryId

# Get results
$results = Receive-GraphAuditXResults -QueryId $queryId
$results | Export-Csv "AuditResults.csv" -NoTypeInformation
```

---

## 🔧 Troubleshooting

| Issue | Solution |
|---|---|
| **"Cannot load assembly PresentationFramework"** | Use **Windows PowerShell 5.1** (`powershell.exe`), NOT PowerShell 7 (`pwsh.exe`) |
| **Window doesn't appear** | Run: `Set-ExecutionPolicy Bypass -Scope CurrentUser` |
| **"Module not found" error** | Verify paths in lines 3, 11, 12 of GraphAuditX-WPF.ps1 |
| **CSV upload shows no data** | CSV must have columns: `RecordType`, `CreationDate`, `Operations`, `AuditData` |
| **HAR shows "No requests found"** | HAR must contain requests to `purview.microsoft.com/apiproxy/adtsch/AuditLogSearch` |
| **Graph query returns 401/403** | Ensure admin consent is granted and secret hasn't expired |
| **"Install-Module Microsoft.Graph" fails** | Run PowerShell as Admin, or add `-Scope CurrentUser` |

---

## ✅ Validation Checklist

After setup, verify:

- [ ] App launches without errors (WPF window appears)
- [ ] Workload dropdown populates (144 workloads)
- [ ] Selecting a workload populates Operations dropdown
- [ ] "Upload CSV" loads data into the grid
- [ ] "Analyze" shows report with dates, top operations, record types
- [ ] Three query buttons generate valid .ps1 scripts
- [ ] "Upload HAR" analyzes a .har file correctly
- [ ] Report text is selectable and Ctrl+C works
- [ ] "Copy All" copies report to clipboard

---

## 📁 Test Files

To validate without a live tenant:

1. **Sample CSV** — Any Purview Audit export from [purview.microsoft.com/audit](https://purview.microsoft.com/audit)
2. **Sample HAR** — Record browser network trace (F12 → Network → Export HAR) while searching in Purview Audit

---

## 📄 License

Internal Microsoft tool — Microsoft Purview CxE Engineering.

---

*Author: Paresh Nhathalal — Microsoft Purview CxE*  
*GitHub: https://github.com/nparesh/GraphAuditX-UI*