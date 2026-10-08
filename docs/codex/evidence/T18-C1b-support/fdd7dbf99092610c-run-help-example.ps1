[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)
.\WinPDFMerge.ps1 '<REPO>\tests\.work\diagnostics\12ce4ca9b4a443a19968e8f81ceb89ce\79c3a958f98f45eb8bbdfea21aad1ce3\source folder' -OutputFolder '<REPO>\tests\.work\diagnostics\12ce4ca9b4a443a19968e8f81ceb89ce\79c3a958f98f45eb8bbdfea21aad1ce3\named output'
exit $LASTEXITCODE
