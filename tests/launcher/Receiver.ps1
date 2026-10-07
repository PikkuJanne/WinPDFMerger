# Synthetic PowerShell receiver for real cmd/.bat tests, not a PDF application.
[CmdletBinding()]
param(
    [string]$SourceFolder,
    [Parameter(ValueFromRemainingArguments=$true)][string[]]$ExtraArguments = @()
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$capture = [Environment]::GetEnvironmentVariable('WINPDFMERGER_LAUNCHER_CAPTURE', 'Process')
$requestedExit = [int][Environment]::GetEnvironmentVariable('WINPDFMERGER_LAUNCHER_EXIT', 'Process')
if (-not $capture -or $requestedExit -lt 0 -or $requestedExit -gt 255) { throw 'Missing test receiver configuration.' }
$policies = [ordered]@{}
foreach ($item in Get-ExecutionPolicy -List) { $policies[$item.Scope.ToString()] = $item.ExecutionPolicy.ToString() }
$record = [ordered]@{
    source_folder = $SourceFolder
    source_exists = [IO.Directory]::Exists($SourceFolder)
    extra_arguments = @($ExtraArguments)
    command_line_arguments = @([Environment]::GetCommandLineArgs())
    script_path = $PSCommandPath
    shell_version = $PSVersionTable.PSVersion.ToString()
    shell_edition = $PSVersionTable.PSEdition
    process_64_bit = [Environment]::Is64BitProcess
    process_id = $PID
    execution_policy = (Get-ExecutionPolicy).ToString()
    execution_policy_scopes = $policies
    requested_exit = $requestedExit
}
[IO.File]::WriteAllText($capture, ($record | ConvertTo-Json -Depth 5), (New-Object Text.UTF8Encoding($false)))
Write-Host 'T05 synthetic receiver completed.'
exit $requestedExit
