param([string]$Entry,[string]$Source,[string]$Output,[string]$Culture='en-US',[switch]$SkipEmail,[switch]$Barrier)
$ErrorActionPreference = 'Stop'
# Match the test process reader's explicit UTF8 decoding. This applies only to
# the owned test child; the user's console and application sources are unchanged.
[Console]::OutputEncoding = New-Object Text.UTF8Encoding($false)
$cultureInfo = [Globalization.CultureInfo]::GetCultureInfo($Culture)
[Threading.Thread]::CurrentThread.CurrentCulture = $cultureInfo
[Threading.Thread]::CurrentThread.CurrentUICulture = $cultureInfo
if ($Barrier) {
    [Console]::WriteLine('CORPUS_READY ' + (ConvertTo-Json -InputObject @{ProcessId=$PID; Culture=$Culture; UtcTicks=[DateTime]::UtcNow.Ticks} -Compress))
    if ([Console]::ReadLine() -cne 'GO') { throw 'The owned launch barrier requires GO.' }
}
$options = @{SourceFolder=$Source; OutputFolder=$Output}
if ($SkipEmail) { $options.SkipEmail = $true }
[Console]::WriteLine('CORPUS_ENTRY_START ' + [DateTime]::UtcNow.Ticks)
& $Entry @options
$entryExitCode = $LASTEXITCODE
[Console]::WriteLine('CORPUS_ENTRY_END ' + [DateTime]::UtcNow.Ticks)
exit $entryExitCode