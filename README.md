Folder	Purpose
Module\GraphAuditX	actual installable module
Public	exported cmdlets
Private	internal helpers (not exposed)
Dev	scratch scripts, GUI, experiments

How to install and use GraphAuditX PowerShell Module which leverages MS Graph for Purview Audit

Import-Module GraphAuditX -Force
$secret = Read-Host "Client Secret" -AsSecureString
Connect-GraphAuditX -TenantId "xxx" -ClientId "xxx" -ClientSecret $secret
$id = GraphAuditX -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date)
Get-GraphAuditXStatus -QueryId $id
Receive-GraphAuditXResults -QueryId $id