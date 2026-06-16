function Connect-GraphAuditX {

    # If already connected, reuse session
    if (Get-MgContext) {
        return
    }

    Write-Host "Opening browser for login..." -ForegroundColor Cyan

    # Delegated auth (browser + MFA)
    Connect-MgGraph -Scopes "AuditLog.Read.All","Directory.Read.All"

    $ctx = Get-MgContext

    if (-not $ctx) {
        throw "Authentication failed"
    }
}