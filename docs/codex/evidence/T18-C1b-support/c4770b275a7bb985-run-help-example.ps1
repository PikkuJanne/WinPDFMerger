[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)
.\WinPDFMerge.ps1 '<REPO>\tests\.work\diagnostics\12ce4ca9b4a443a19968e8f81ceb89ce\aba2a4e6821141c1a02e57f2f3af0e58\source folder'
exit $LASTEXITCODE
