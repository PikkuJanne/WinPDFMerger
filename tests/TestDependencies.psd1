@{
    # Explicit import in tools/test/Invoke-Tests.ps1; never auto-installs modules.
    PesterVersion = '6.2.0'
    PowerShellMinimum = '5.1'
    PesterPowerShellCoreMinimum = '7.4'
    # Explicit development reference pins; no automatic download or install.
    ReferencePowerShellCoreVersion = '7.6.6'
    PSScriptAnalyzerVersion = '1.25.0'
    # Development only. Existing T19 runtime and fresh T21 workspace bundle
    # 26.1007.11041 provide the same Python/package versions with different exe
    # bytes. Both exact observed executables are accepted; no automatic install.
    DevelopmentPythonVersion = '3.12.14'
    DevelopmentPythonSHA256 = @(
        'dd5f8d19f6755d6491ee7c4bef2fe35ddd521334cc3ca3ed8fc93ebcadf135d0'
        '10d845f50a2af64e3500bb2fcb348b5bc98a75d8ddada63e45ba1da6a1fc79d1'
    )
}
