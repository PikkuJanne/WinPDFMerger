[CmdletBinding()]param([string]$Configuration)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$config=Get-Content -LiteralPath $Configuration -Raw | ConvertFrom-Json
[Threading.Thread]::CurrentThread.CurrentCulture=[Globalization.CultureInfo]::GetCultureInfo('de-DE')
[Threading.Thread]::CurrentThread.CurrentUICulture=[Globalization.CultureInfo]::GetCultureInfo('de-DE')
try { & $config.Entry -SourceFolder $config.Source -OutputFolder $config.Output -SkipEmail:($config.Mode -eq 'skipped'); exit $LASTEXITCODE }
catch { [Console]::Error.WriteLine($_.Exception.Message); exit 1 }