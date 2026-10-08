[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)
.\WinPDFMerge.ps1 '<REPO>\tests\.work\diagnostics\98c49c3c31d845cf812f617d355b2667\55c61f1a013c497b9b676b0887238ee1\source folder' -EmailPreset ebook
exit $LASTEXITCODE
