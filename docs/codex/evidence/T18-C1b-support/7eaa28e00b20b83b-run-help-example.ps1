[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)
.\WinPDFMerge.ps1 '<REPO>\tests\.work\diagnostics\276b72e7f066414b8eff703e28391f0b\7961e5165fa9432ea92d6465b307e84a\source folder' -OutputFolder '<REPO>\tests\.work\diagnostics\276b72e7f066414b8eff703e28391f0b\7961e5165fa9432ea92d6465b307e84a\named output'
exit $LASTEXITCODE
