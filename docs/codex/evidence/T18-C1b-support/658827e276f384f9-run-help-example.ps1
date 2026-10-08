[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)
.\WinPDFMerge.ps1 '<REPO>\tests\.work\diagnostics\34a9855893114012a6901dfcfcd027e4\20c9512c05c7404297f7610662c81b56\source folder' -EmailPreset ebook
exit $LASTEXITCODE
