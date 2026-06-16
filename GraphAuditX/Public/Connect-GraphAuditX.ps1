function Connect-GraphAuditX {

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][securestring]$ClientSecret
    )

    Write-Host "🔐 Connecting to Microsoft Graph..." -ForegroundColor Cyan

    $plainSecret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($ClientSecret)
    )

    $token = Get-GraphAuditXToken `
        -TenantId $TenantId `
        -ClientId $ClientId `
        -ClientSecret $plainSecret

    $script:GraphAuditXAuth = @{
        Token  = $token
        Expiry = (Get-Date).AddHours(1)
    }

    Write-Host "✅ Connected successfully" -ForegroundColor Green
}