[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)
.\WinPDFMerge.ps1 '<REPO>\tests\.work\diagnostics\12ce4ca9b4a443a19968e8f81ceb89ce\d6b3156275144bdc9fc1d16e0e61821a\source folder' -EmailPreset ebook
exit $LASTEXITCODE
