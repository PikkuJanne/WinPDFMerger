$win = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
$channel = if (Test-Path -LiteralPath 'HKLM:\SOFTWARE\Microsoft\WindowsSelfHost\Applicability') {
    Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\WindowsSelfHost\Applicability' |
        Select-Object BranchName, ContentType, Ring
} else { $null }
[pscustomobject]@{
    observed_at_utc = [DateTime]::UtcNow.ToString('o')
    commit = (git rev-parse HEAD)
    dirty_worktree = [bool]@(git status --porcelain=v1)
    edition = $win.EditionID
    display_version = $win.DisplayVersion
    current_build = $win.CurrentBuildNumber
    ubr = $win.UBR
    full_build = ($win.CurrentBuildNumber + '.' + $win.UBR)
    os_version = [Environment]::OSVersion.Version.ToString()
    process_64_bit = [Environment]::Is64BitProcess
    is_administrator = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    shell_version = $PSVersionTable.PSVersion.ToString()
    shell_edition = $PSVersionTable.PSEdition
    channel_registry = $channel
    policy = @(Get-ExecutionPolicy -List | ForEach-Object {
        [pscustomobject]@{scope = $_.Scope.ToString(); policy = $_.ExecutionPolicy.ToString()}
    })
    observation_class = 'read-only inventory; no Explorer or PDF-viewer observations'
    insider_enrollment = 'unobserved; null registry values are not proof'
} | ConvertTo-Json -Depth 6
