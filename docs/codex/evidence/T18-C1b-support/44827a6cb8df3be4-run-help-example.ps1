[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)
.\WinPDFMerge.ps1 '<REPO>\tests\.work\diagnostics\98c49c3c31d845cf812f617d355b2667\a246f19af649415a8ad22eadff254b05\source folder' -SkipEmail
exit $LASTEXITCODE
