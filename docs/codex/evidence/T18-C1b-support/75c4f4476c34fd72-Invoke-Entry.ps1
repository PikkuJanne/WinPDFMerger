[CmdletBinding()]param([string]$Configuration)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$config=Get-Content -LiteralPath $Configuration -Raw | ConvertFrom-Json
[Threading.Thread]::CurrentThread.CurrentCulture=[Globalization.CultureInfo]::GetCultureInfo('de-DE')
[Threading.Thread]::CurrentThread.CurrentUICulture=[Globalization.CultureInfo]::GetCultureInfo('de-DE')
$named=@{}; foreach($property in $config.Named.PSObject.Properties){$named[$property.Name]=$property.Value}
try { & $config.Entry @named; exit $LASTEXITCODE }
catch {
 $receipt=Get-Content -LiteralPath $config.Receipt -Raw | ConvertFrom-Json; $receipt.BindingError=$_.Exception.Message
 [IO.File]::WriteAllText($config.Receipt,($receipt | ConvertTo-Json -Depth 24),(New-Object Text.UTF8Encoding($false)))
 [Console]::Error.WriteLine($_.Exception.Message); exit 1
}