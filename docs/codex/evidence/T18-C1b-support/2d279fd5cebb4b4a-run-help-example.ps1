[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)
.\WinPDFMerge.ps1 '<REPO>\tests\.work\diagnostics\98c49c3c31d845cf812f617d355b2667\b65152d147e34cdf880f045e3e278f89\source folder'
exit $LASTEXITCODE
