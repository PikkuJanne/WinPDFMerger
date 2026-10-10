[CmdletBinding()]
param([string]$SourceRoot,[string]$OutputDirectory,[string]$SourceCommit)
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
& (Join-Path $SourceRoot 'tools/release/Build-Release.ps1') -SourceCommit $SourceCommit -OutputDirectory $OutputDirectory -RepositoryRoot $SourceRoot | ConvertTo-Json -Depth 12
