[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)
.\WinPDFMerge.ps1 '<REPO>\tests\.work\diagnostics\34a9855893114012a6901dfcfcd027e4\bcf31066b35c415d8f725028d2387072\source folder'
exit $LASTEXITCODE
