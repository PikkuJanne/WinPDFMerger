[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)
.\WinPDFMerge.ps1 '<REPO>\tests\.work\diagnostics\98c49c3c31d845cf812f617d355b2667\9a98ecd4ca2142e2ae6a325e72538822\source folder' -OutputFolder '<REPO>\tests\.work\diagnostics\98c49c3c31d845cf812f617d355b2667\9a98ecd4ca2142e2ae6a325e72538822\named output'
exit $LASTEXITCODE
