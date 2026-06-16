@{
    RootModule        = 'GraphAuditX.psm1'
    ModuleVersion     = '1.1.0'
    GUID              = '3f8b2c6e-1a2d-4f6b-9c1e-7d3a5b8e9f01'

    Author            = 'Paresh Nhathalal'
    CompanyName       = 'Microsoft'
    Copyright         = '(c) Paresh Nhathalal. All rights reserved.'

    Description       = 'GraphAuditX: Microsoft Graph-powered replacement for Search-UnifiedAuditLog with CLI parity.'

    PowerShellVersion = '5.1'

    FunctionsToExport = @('*')
    CmdletsToExport   = @()
    VariablesToExport = '*'

    # ✅ Alias exported automatically on Import-Module
    AliasesToExport   = @('Search-UnifiedAuditLog')

    PrivateData = @{
        PSData = @{
            Tags       = @('MicrosoftGraph','Purview','Audit','Security','UnifiedAuditLog')
            ProjectUri = 'https://github.com/nparesh/GraphAuditX-PowerShell-Module'
        }
    }
}