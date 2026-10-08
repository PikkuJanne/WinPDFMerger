[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)
.\WinPDFMerge.ps1 '<REPO>\tests\.work\diagnostics\12ce4ca9b4a443a19968e8f81ceb89ce\5341dddcbed74afba67b202cfff92200\source folder' -SkipEmail
exit $LASTEXITCODE
