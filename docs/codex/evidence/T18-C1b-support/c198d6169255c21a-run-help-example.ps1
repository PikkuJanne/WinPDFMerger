[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)
.\WinPDFMerge.ps1 '<REPO>\tests\.work\diagnostics\34a9855893114012a6901dfcfcd027e4\919526c17aac4c10968864e8fda111c0\source folder' -OutputFolder '<REPO>\tests\.work\diagnostics\34a9855893114012a6901dfcfcd027e4\919526c17aac4c10968864e8fda111c0\named output'
exit $LASTEXITCODE
