@{
    # Script module or binary module file associated with this manifest.
    ModuleVersion = '1.2'
    GUID = 'xxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxx'  # Generate a GUID using https://www.guidgen.com/ if needed
    Author = 'Paresh Nhathalal'
    CompanyName = 'YourCompanyName'
    Copyright = '© YourCompanyName 2026'
    Description = 'A PowerShell module to query Microsoft Purview audit data via Microsoft Graph API'
    PowerShellVersion = '5.1'  # Make sure this is compatible with most PowerShell versions
    ModulesRequired = @()  # Leave empty unless you depend on other modules
    FunctionsToExport = @(
        'Connect-GraphAuditX',
        'GraphAuditX',
        'Get-GraphAuditXStatus',
        'Receive-GraphAuditXResults'
    )
    # Specify any dependencies your module may require (leave empty if no dependencies)
    RequiredModules = @()
    # URL to your project on GitHub
    ProjectUri = 'https://github.com/nparesh/GraphAuditX-PowerShell-Module'
    # Optional, specify your license
    LicenseUri = 'https://opensource.org/licenses/MIT'
    # Copyright info
    Copyright = '© YourCompanyName'
}