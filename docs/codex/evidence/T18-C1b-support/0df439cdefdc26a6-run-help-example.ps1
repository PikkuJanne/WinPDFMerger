[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)
.\WinPDFMerge.ps1 '<REPO>\tests\.work\diagnostics\276b72e7f066414b8eff703e28391f0b\b0974400b3af409787184f2bc1bed8ab\source folder'
exit $LASTEXITCODE
