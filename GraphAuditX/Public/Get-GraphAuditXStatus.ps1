function Get-GraphAuditXStatus {
<#
.SYNOPSIS
Gets the status of a GraphAuditX audit query

.VERSION
1.2
#>

    param(
        [Parameter(Mandatory)]
        [string]$QueryId
    )

    # ===== CONNECTION VALIDATION =====
    try {
        $token = Get-GraphAuditXToken
        if (-not $token) { throw "No token" }
    }
    catch {
        throw "❌ Not connected. Run Connect-GraphAuditX first."
    }

    $uri = "https://graph.microsoft.com/beta/security/auditLog/queries/$QueryId"

    $response = Invoke-GraphAuditXRequest -Method GET -Uri $uri

    if ($response -and $response.status) {
        Write-Host "Status: $($response.status)" -ForegroundColor Yellow
        return $response.status
    }
    else {
        throw "❌ Failed to retrieve query status"
    }
}