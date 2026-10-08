[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)
.\WinPDFMerge.ps1 '<REPO>\tests\.work\diagnostics\34a9855893114012a6901dfcfcd027e4\9abb275c91194437963535f3b35c86f8\source folder' -SkipEmail
exit $LASTEXITCODE
