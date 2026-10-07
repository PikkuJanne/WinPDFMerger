@{
    # Explicit import in tools/test/Invoke-Tests.ps1; never auto-installs modules.
    PesterVersion = '6.2.0'
    PowerShellMinimum = '5.1'
    PesterPowerShellCoreMinimum = '7.4'
    # Explicit development reference pins; no automatic download or install.
    ReferencePowerShellCoreVersion = '7.6.6'
    PSScriptAnalyzerVersion = '1.25.0'
}
