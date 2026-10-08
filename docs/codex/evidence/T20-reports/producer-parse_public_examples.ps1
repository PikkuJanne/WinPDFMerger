# Read-only T20 syntax review. Parses fenced text; never invokes a doc example.
$t20ParserPaths=@('README.md','docs/USAGE.md','docs/TROUBLESHOOTING.md','docs/DEPENDENCIES.md','SECURITY.md')
$t20ParserBlocks=New-Object 'System.Collections.Generic.List[object]'
foreach($t20ParserPath in $t20ParserPaths) {
    $t20ParserText=[IO.File]::ReadAllText((Join-Path (Get-Location).Path $t20ParserPath))
    foreach($t20ParserFence in [regex]::Matches($t20ParserText,'(?ms)^```powershell\r?\n(.*?)^```')) {
        $t20ParserTokens=$null; $t20ParserErrors=$null
        $t20ParserAst=[Management.Automation.Language.Parser]::ParseInput($t20ParserFence.Groups[1].Value,[ref]$t20ParserTokens,[ref]$t20ParserErrors)
        $t20ParserCommands=@($t20ParserAst.FindAll({param($t20ParserNode) $t20ParserNode -is [Management.Automation.Language.CommandAst]},$true) | ForEach-Object {$_.Extent.Text})
        $t20ParserBlocks.Add([pscustomobject]@{Path=$t20ParserPath;Line=([regex]::Matches($t20ParserText.Substring(0,$t20ParserFence.Index),'\n').Count+1);Commands=$t20ParserCommands;Errors=@($t20ParserErrors | ForEach-Object {$_.Message});Result=$(if(@($t20ParserErrors).Count -eq 0){'pass'}else{'fail'})})
    }
}
$t20ParserRecord=[ordered]@{Schema='t20.public-example-syntax-review.v1';Utc=[datetime]::UtcNow.ToString('o');Scope='actual PowerShell AST parse of public fenced examples; no application/example/native invocation';Shell=$PSVersionTable.PSVersion.ToString();Edition=$PSVersionTable.PSEdition;ProducerSHA256=(Get-FileHash -LiteralPath $PSCommandPath -Algorithm SHA256).Hash.ToLowerInvariant();SourceBindings=@($t20ParserPaths | ForEach-Object {[pscustomobject]@{Path=$_;SHA256=(Get-FileHash -LiteralPath $_ -Algorithm SHA256).Hash.ToLowerInvariant()}});Blocks=$t20ParserBlocks.ToArray();BlockCount=$t20ParserBlocks.Count;CommandCount=(@($t20ParserBlocks.ToArray() | ForEach-Object {$_.Commands})).Count;ApplicationInvoked=$false;NativeEngineInvoked=$false;Result=$(if(@($t20ParserBlocks.ToArray() | Where-Object Result -ne pass).Count -eq 0){'pass'}else{'fail'})}
$t20ParserDestination='tests/.work/T20-user-docs-final-review/ast-review.json'
[IO.File]::WriteAllText($t20ParserDestination,($t20ParserRecord | ConvertTo-Json -Depth 10),(New-Object Text.UTF8Encoding($false)))
[pscustomobject]@{Result=$t20ParserRecord.Result;Blocks=$t20ParserRecord.BlockCount;Commands=$t20ParserRecord.CommandCount;Shell=$t20ParserRecord.Shell;Report=$t20ParserDestination;SHA256=(Get-FileHash -LiteralPath $t20ParserDestination -Algorithm SHA256).Hash.ToLowerInvariant()} | ConvertTo-Json -Compress
if($t20ParserRecord.Result -ne 'pass') {exit 1}
