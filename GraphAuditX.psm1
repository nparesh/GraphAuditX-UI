# This loads the functions from Public and Private folders into the module
Get-ChildItem "$PSScriptRoot\Private\*.ps1" | ForEach-Object { . $_.FullName }
Get-ChildItem "$PSScriptRoot\Public\*.ps1"  | ForEach-Object { . $_.FullName }

# Alias
Set-Alias -Name Search-UnifiedAuditLog -Value GraphAuditX