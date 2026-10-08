[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)
.\WinPDFMerge.ps1 '<REPO>\tests\.work\diagnostics\276b72e7f066414b8eff703e28391f0b\06850bf90eee44a28935664b0ed5ba1f\source folder' -SkipEmail
exit $LASTEXITCODE
