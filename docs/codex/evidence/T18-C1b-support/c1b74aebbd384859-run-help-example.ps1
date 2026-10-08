[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)
.\WinPDFMerge.ps1 '<REPO>\tests\.work\diagnostics\276b72e7f066414b8eff703e28391f0b\2a30931666d145a8a65275c539ed768f\source folder' -EmailPreset ebook
exit $LASTEXITCODE
