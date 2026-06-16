# GraphAuditX - Setup Guide

> **GraphAuditX Pro** is a WPF-based PowerShell tool for querying Microsoft 365 Purview Audit Logs via Microsoft Graph, O365 Management API, and Search-UnifiedAuditLog. It includes CSV analysis, HAR file parsing, and query generation.

---

## Prerequisites

| Requirement | Details |
|---|---|
| **OS** | Windows 10 / 11 / Server 2016+ |
| **PowerShell** | Windows PowerShell 5.1 (ships with Windows) |
| **.NET Framework** | 4.7.2+ (for WPF / PresentationFramework) |
| **Microsoft.Graph module** | Optional — only needed for live Graph API queries |
| **Entra App Registration** | Optional — only needed for live audit log queries |

---

## Step 1: Download the Tool

Clone or copy the entire `GraphAuditX-PowerShell-Module` folder to your machine. The recommended location is:

```
C:\GraphAuditX-PowerShell-Module\
```

### Folder Structure

```
GraphAuditX-PowerShell-Module\
├── GraphAuditX-WPF.ps1              ← Main WPF GUI application
├── AuditOperationsGateConfig.csv    ← Workload/Operations config data
├── GraphAuditX.psd1                 ← Module manifest
├── GraphAuditX.psm1                 ← Module loader
├── RecordTypeResolver.cs            ← Record type enum resolver
├── GraphAuditX\                     ← Core module
│   ├── GraphAuditX.psd1
│   ├── GraphAuditX.psm1
│   ├── Public\                      ← Exported functions
│   │   ├── Connect-GraphAuditX.ps1
│   │   ├── GraphAuditX.ps1
│   │   ├── Get-GraphAuditXStatus.ps1
│   │   └── Receive-GraphAuditXResults.ps1
│   └── Private\                     ← Internal helpers
│       ├── Connect-GraphAuditX.ps1
│       ├── Get-GraphAuditXToken.ps1
│       └── Invoke-GraphAuditXRequest.ps1
├── M365Api\                         ← O365 Management API support
│   └── M365-ManagementAPI.ps1
└── Assets\                          ← Icons/images (optional)
```

---

## Step 2: Update Hardcoded Paths

Open `GraphAuditX-WPF.ps1` and update the **3 paths at the top** (lines 3, 11, 12) to match your install location:

```powershell
# Line 3 - M365 Management API helper
. "C:\GraphAuditX-PowerShell-Module\M365Api\M365-ManagementAPI.ps1"

# Line 11 - Module import
Import-Module "C:\GraphAuditX-PowerShell-Module\GraphAuditX\GraphAuditX.psd1" -Force

# Line 12 - Config CSV
$configPath = "C:\GraphAuditX-PowerShell-Module\AuditOperationsGateConfig.csv"
```

**Replace** `C:\GraphAuditX-PowerShell-Module` with your actual folder path.

For example, if you placed it in `D:\Tools\GraphAuditX-PowerShell-Module`:

```powershell
. "D:\Tools\GraphAuditX-PowerShell-Module\M365Api\M365-ManagementAPI.ps1"
Import-Module "D:\Tools\GraphAuditX-PowerShell-Module\GraphAuditX\GraphAuditX.psd1" -Force
$configPath = "D:\Tools\GraphAuditX-PowerShell-Module\AuditOperationsGateConfig.csv"
```

---

## Step 3: Set Execution Policy (if needed)

If you haven't already, allow script execution:

```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope CurrentUser
```

---

## Step 4: Launch the Tool

```powershell
powershell.exe -ExecutionPolicy Bypass -File "C:\GraphAuditX-PowerShell-Module\GraphAuditX-WPF.ps1"
```

Or from within PowerShell:

```powershell
& "C:\GraphAuditX-PowerShell-Module\GraphAuditX-WPF.ps1"
```

The **GraphAuditX Pro** WPF window should appear.

---

## Step 5: (Optional) Set Up Entra App for Live Queries

To run **live audit log queries** against your tenant, you need an Entra (Azure AD) App Registration:

1. Go to [Azure Portal → Entra ID → App registrations](https://portal.azure.com/#view/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/~/RegisteredApps)
2. Click **New registration**
   - Name: `GraphAuditX`
   - Supported account types: Single tenant
3. Under **API permissions**, add:
   - `Microsoft Graph` → Application permissions:
     - `AuditLogsQuery.Read.All`
     - `AuditLog.Read.All`
   - `Office 365 Management APIs` → Application permissions:
     - `ActivityFeed.Read`
     - `ActivityFeed.ReadDlp`
4. Click **Grant admin consent**
5. Under **Certificates & secrets**, create a **Client Secret** and note the value
6. Note your **Tenant ID** and **Client (Application) ID**

Enter these in the tool's connection fields when running live queries.

---

## Features Overview

### 🔍 Live Audit Query
- Select workloads, record types, operations, date range, and users
- Queries via Microsoft Graph API (beta/security/auditLog/queries)
- Results displayed in filterable DataGrid

### 📂 Upload CSV
- Upload a previously exported Purview Audit CSV
- Loads records into the grid for filtering
- **Analyze** generates a report: top operations, workloads, record types, user activity
- **Query Generation**: Get MS Graph Query / O365 API Query / PowerShell Query
  - Each generates a ready-to-run .ps1 script pre-filled with the CSV's parameters

### 🔍 Upload HAR
- Upload a browser HAR file (network trace from Purview portal)
- Extracts all Purview `AuditLogSearch` API calls
- Shows Query IDs, search parameters, date ranges, operations, users
- Useful for reproducing a search someone performed in the portal

### 📋 Export
- Export results to CSV
- Open in Excel directly

---

## Troubleshooting

| Issue | Solution |
|---|---|
| **"Cannot load assembly PresentationFramework"** | Ensure you're running **Windows PowerShell 5.1** (not PowerShell 7/Core). Run: `powershell.exe` not `pwsh.exe` |
| **Window doesn't appear** | Check `Set-ExecutionPolicy Bypass -Scope CurrentUser` |
| **"Module not found" error** | Verify the paths on lines 3, 11, 12 match your install folder |
| **CSV upload shows no data** | Ensure the CSV has columns: `RecordType`, `CreationDate`, `Operations`, `AuditData` |
| **HAR shows "No requests found"** | The HAR file must contain requests to `purview.microsoft.com/apiproxy/adtsch/AuditLogSearch` |
| **Graph query fails at runtime** | Install the Microsoft.Graph module: `Install-Module Microsoft.Graph -Scope CurrentUser` |

---

## Quick Validation Checklist

After setup, verify these work:

- [ ] App launches without errors
- [ ] Workload dropdown populates with items (dark black text)
- [ ] Selecting a workload populates Operations dropdown
- [ ] "Upload CSV" opens file picker and loads a CSV into the grid
- [ ] "Upload CSV" → Analyze shows report with dates, top operations, record types
- [ ] Three query buttons each open Notepad with a valid .ps1 script
- [ ] "Upload HAR" opens file picker and analyzes a .har file
- [ ] Report text is selectable (click & drag) and Ctrl+C works
- [ ] "Copy All" button copies report to clipboard

---

## Minimum Test Files

To validate without a live tenant connection:

1. **Sample CSV** — Any Purview Audit export CSV with columns: `RecordType`, `CreationDate`, `UserIds`, `Operations`, `AuditData`
2. **Sample HAR** — Record a browser network trace while searching in [Purview Compliance Portal](https://purview.microsoft.com) → Audit

---

*Author: Paresh Nhathalal | Microsoft CxE EEE*
*Version: 6.0 | Last Updated: June 2026*
