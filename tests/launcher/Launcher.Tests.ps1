BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
    . (Join-Path $repo 'tests/TestSupport.ps1')
    . (Join-Path $PSScriptRoot 'TestSupport.ps1')
    $batch = Join-Path $repo 'WinPDFMerge.bat'
    $receiver = Join-Path $PSScriptRoot 'Receiver.ps1'
    $parentPath = [Environment]::GetEnvironmentVariable('PATH', 'Process')
    $parentModulePath = [Environment]::GetEnvironmentVariable('PSModulePath', 'Process')
    $parentErrorLevel = [Environment]::GetEnvironmentVariable('ERRORLEVEL', 'Process')
    $parentPolicies = @(Get-ExecutionPolicy -List | ForEach-Object { $_.Scope.ToString() + '=' + $_.ExecutionPolicy.ToString() }) -join ';'

    function New-LauncherCase([string]$InstallName = 'install', [string]$SourceName = 'source', [int]$ExitCode = 0) {
        $run = Join-Path $repo ('tests/.work/launcher/' + [Guid]::NewGuid().ToString('N'))
        $install = Join-Path $run $InstallName
        $source = Join-Path $run $SourceName
        [void][IO.Directory]::CreateDirectory($install)
        [void][IO.Directory]::CreateDirectory($source)
        [IO.File]::Copy($batch, (Join-Path $install 'WinPDFMerge.bat'))
        [IO.File]::Copy($receiver, (Join-Path $install 'WinPDFMerge.ps1'))
        $capture = Join-Path $run 'receiver.json'
        [pscustomobject]@{
            Root = $run
            Install = $install
            Source = $source
            Batch = (Join-Path $install 'WinPDFMerge.bat')
            Script = (Join-Path $install 'WinPDFMerge.ps1')
            Capture = $capture
            Environment = @{
                WINPDFMERGER_LAUNCHER_CAPTURE = $capture
                WINPDFMERGER_LAUNCHER_EXIT = [string]$ExitCode
                T05_SENTINEL = 'EXPANDED'
                # Also applies to the direct-PowerShell alternative below.
                PSModulePath = (Join-Path $env:SystemRoot 'System32/WindowsPowerShell/v1.0/Modules')
            }
        }
    }

    function Assert-LauncherReceiver($Case, [string]$ExpectedSource = $Case.Source) {
        [IO.File]::Exists($Case.Capture) | Should -BeTrue
        $received = Get-Content -LiteralPath $Case.Capture -Raw | ConvertFrom-Json
        $received.source_folder | Should -BeExactly $ExpectedSource
        $received.source_exists | Should -BeTrue
        @($received.extra_arguments).Count | Should -Be 0
        $received.shell_edition | Should -BeExactly 'Desktop'
        $received.shell_version | Should -Match '^5\.1\.'
        $received.process_64_bit | Should -BeTrue
        $received.command_line_arguments | Should -Contain '-NoProfile'
        $received.command_line_arguments | Should -Contain '-ExecutionPolicy'
        $received.command_line_arguments | Should -Contain 'Bypass'
        $received.command_line_arguments | Should -Contain '-File'
        $received.command_line_arguments | Should -Contain '-SourceFolder'
        $received.execution_policy | Should -BeExactly 'Bypass'
        $received.execution_policy_scopes.MachinePolicy | Should -BeExactly 'Undefined'
        $received.execution_policy_scopes.UserPolicy | Should -BeExactly 'Undefined'
        [Environment]::GetEnvironmentVariable('PATH', 'Process') | Should -BeExactly $parentPath
        [Environment]::GetEnvironmentVariable('PSModulePath', 'Process') | Should -BeExactly $parentModulePath
        (@(Get-ExecutionPolicy -List | ForEach-Object { $_.Scope.ToString() + '=' + $_.ExecutionPolicy.ToString() }) -join ';') | Should -BeExactly $parentPolicies
        return $received
    }
}

Describe 'AC009: actual cmd batch path safety with a synthetic PowerShell receiver' {
    It 'preserves <Label> in installation and source paths' -TestCases @(
        @{ Label = 'spaces'; Suffix = 'two words' },
        @{ Label = 'exclamation marks'; Suffix = '!keep!' },
        @{ Label = 'ampersands'; Suffix = '& group' },
        @{ Label = 'parentheses'; Suffix = '(group)' },
        @{ Label = 'brackets'; Suffix = '[group]' },
        @{ Label = 'combined metacharacters'; Suffix = 'space ! & (group) [set]' }
    ) {
        param($Label, $Suffix)
        $case = New-LauncherCase ('install ' + $Suffix) ('source ' + $Suffix)
        $result = Invoke-LauncherCommand -BatchPath $case.Batch -SourceArguments @($case.Source) -ChildEnvironment $case.Environment
        $result.ExitCode | Should -Be 0 -Because ($result.Stdout + $result.Stderr)
        [void](Assert-LauncherReceiver $case)
    }

    It 'does not execute an ampersand-separated command embedded in an installation path' {
        $case = New-LauncherCase 'install & echo T05_UNINTENDED_INSTALL' 'source'
        $result = Invoke-LauncherCommand -BatchPath $case.Batch -SourceArguments @($case.Source) -ChildEnvironment $case.Environment
        $result.ExitCode | Should -Be 0 -Because ($result.Stdout + $result.Stderr)
        $result.Stdout | Should -Not -Match '(?m)^T05_UNINTENDED_INSTALL'
        [void](Assert-LauncherReceiver $case)
    }

    It 'quotes an invalid source diagnostic without executing its embedded echo command' {
        $case = New-LauncherCase
        $missing = Join-Path $case.Root 'missing & echo T05_UNINTENDED_SOURCE'
        $result = Invoke-LauncherCommand -BatchPath $case.Batch -SourceArguments @($missing) -ChildEnvironment $case.Environment
        $result.ExitCode | Should -Be 1
        $result.Stdout | Should -Match 'not a folder|not a directory'
        $result.Stdout | Should -Not -Match '(?m)^T05_UNINTENDED_SOURCE'
        [IO.File]::Exists($case.Capture) | Should -BeFalse
    }

    It 'characterizes paired percent expansion in the outer cmd command before the batch receives input' {
        $case = New-LauncherCase 'install' 'source%T05_SENTINEL%'
        $expanded = Join-Path $case.Root 'sourceEXPANDED'
        [void][IO.Directory]::CreateDirectory($expanded)
        $result = Invoke-LauncherCommand -BatchPath $case.Batch -SourceArguments @($case.Source) -ChildEnvironment $case.Environment
        $result.ExitCode | Should -Be 0 -Because ($result.Stdout + $result.Stderr)
        [void](Assert-LauncherReceiver $case $expanded)
        [IO.Directory]::Exists($case.Source) | Should -BeTrue
        Write-Host 'T05 boundary observation: cmd expanded source%T05_SENTINEL% to sourceEXPANDED before the launcher received it.'
    }

    It 'characterizes paired percent expansion in the installation path at the outer cmd boundary' {
        $case = New-LauncherCase 'install%T05_SENTINEL%' 'source'
        $expandedInstall = Join-Path $case.Root 'installEXPANDED'
        [void][IO.Directory]::CreateDirectory($expandedInstall)
        [IO.File]::Copy($case.Batch, (Join-Path $expandedInstall 'WinPDFMerge.bat'))
        [IO.File]::Copy($case.Script, (Join-Path $expandedInstall 'WinPDFMerge.ps1'))
        $result = Invoke-LauncherCommand -BatchPath $case.Batch -SourceArguments @($case.Source) -ChildEnvironment $case.Environment
        $result.ExitCode | Should -Be 0 -Because ($result.Stdout + $result.Stderr)
        $received = Assert-LauncherReceiver $case
        $received.script_path | Should -BeExactly (Join-Path $expandedInstall 'WinPDFMerge.ps1')
        [IO.Directory]::Exists($case.Install) | Should -BeTrue
        Write-Host 'T05 boundary observation: cmd selected installEXPANDED when given install%T05_SENTINEL% in its outer command.'
    }

    It 'preserves the literal paired-percent source through the direct PowerShell alternative' {
        $case = New-LauncherCase 'install%T05_SENTINEL%' 'source%T05_SENTINEL%'
        $windowsPowerShell = Join-Path $env:SystemRoot 'System32/WindowsPowerShell/v1.0/powershell.exe'
        $result = Invoke-TestChildProcess -Executable $windowsPowerShell -Arguments @('-NoProfile', '-ExecutionPolicy', 'RemoteSigned', '-File', $case.Script, '-SourceFolder', $case.Source) -ChildEnvironment $case.Environment
        $result.ExitCode | Should -Be 0 -Because ($result.Stdout + $result.Stderr)
        $received = Get-Content -LiteralPath $case.Capture -Raw | ConvertFrom-Json
        $received.source_folder | Should -BeExactly $case.Source
        $received.script_path | Should -BeExactly $case.Script
        $received.source_exists | Should -BeTrue
        $received.execution_policy | Should -BeExactly 'RemoteSigned'
        Write-Host 'T05 boundary alternative: direct powershell.exe -File receives the literal percent-bearing SourceFolder without cmd parsing.'
    }

    It 'passes a trailing source separator without absorbing the closing native quote' {
        $case = New-LauncherCase 'install' 'source'
        $trailing = $case.Source + '\'
        $result = Invoke-LauncherCommand -BatchPath $case.Batch -SourceArguments @($trailing) -ChildEnvironment $case.Environment
        $result.ExitCode | Should -Be 0 -Because ($result.Stdout + $result.Stderr)
        [void](Assert-LauncherReceiver $case $trailing)
    }

    It 'passes multiple trailing source separators without losing the argument boundary' {
        $case = New-LauncherCase
        $trailing = $case.Source + '\\\'
        $result = Invoke-LauncherCommand -BatchPath $case.Batch -SourceArguments @($trailing) -ChildEnvironment $case.Environment
        $result.ExitCode | Should -Be 0 -Because ($result.Stdout + $result.Stderr)
        [void](Assert-LauncherReceiver $case $trailing)
    }

    It 'preserves a filesystem root argument through the trailing-backslash quoting path' {
        $case = New-LauncherCase
        $root = [IO.Path]::GetPathRoot($case.Root)
        $result = Invoke-LauncherCommand -BatchPath $case.Batch -SourceArguments @($root) -ChildEnvironment $case.Environment
        $result.ExitCode | Should -Be 0 -Because ($result.Stdout + $result.Stderr)
        [void](Assert-LauncherReceiver $case $root)
    }
}

Describe 'AC010: actual cmd batch invocation and exit contract with a synthetic receiver' {
    It 'presents exit <Code> and propagates the exact code' -TestCases @(
        @{ Code = 0; Message = '(?i)completed successfully|success' },
        @{ Code = 1; Message = '(?i)failed.*(?:exit code )?1|failure' },
        @{ Code = 2; Message = '(?i)partial success|master.*(?:retained|safe)' },
        @{ Code = 7; Message = '(?i)failed.*(?:exit code )?7|failure' }
    ) {
        param($Code, $Message)
        $case = New-LauncherCase 'install' 'source' $Code
        $result = Invoke-LauncherCommand -BatchPath $case.Batch -SourceArguments @($case.Source) -ChildEnvironment $case.Environment
        $result.ExitCode | Should -Be $Code
        $result.Stdout | Should -Match $Message
        if ($Code -ne 0) { $result.Stdout | Should -Not -Match 'Merge completed successfully' }
        [void](Assert-LauncherReceiver $case)
    }

    It 'retains an actual pause until controlled stdin is supplied' {
        $case = New-LauncherCase
        $result = Invoke-LauncherCommand -BatchPath $case.Batch -SourceArguments @($case.Source) -ChildEnvironment $case.Environment -ObservePause
        $result.ReceiverExitedBeforePauseInput | Should -BeTrue
        $result.AwaitedPauseInput | Should -BeTrue
        $result.ExitCode | Should -Be 0
        [void](Assert-LauncherReceiver $case)
    }

    It 'rejects zero arguments with usage and exit one before invoking the receiver' {
        $case = New-LauncherCase
        $result = Invoke-LauncherCommand -BatchPath $case.Batch -ChildEnvironment $case.Environment
        $result.ExitCode | Should -Be 1
        $result.Stdout | Should -Match '(?i)usage:|drag.and.drop a folder'
        [IO.File]::Exists($case.Capture) | Should -BeFalse
    }

    It 'rejects multiple dropped source paths with usage rather than ignoring the second' {
        $case = New-LauncherCase
        $second = Join-Path $case.Root 'second source'
        [void][IO.Directory]::CreateDirectory($second)
        $result = Invoke-LauncherCommand -BatchPath $case.Batch -SourceArguments @($case.Source, $second) -ChildEnvironment $case.Environment
        $result.ExitCode | Should -Be 1
        $result.Stdout | Should -Match '(?i)usage:|exactly one|one folder'
        [IO.File]::Exists($case.Capture) | Should -BeFalse
    }

    It 'rejects an explicitly empty second argument as an extra argument' {
        $case = New-LauncherCase
        $result = Invoke-LauncherCommand -BatchPath $case.Batch -SourceArguments @($case.Source, '') -ChildEnvironment $case.Environment
        $result.ExitCode | Should -Be 1
        $result.Stdout | Should -Match '(?i)usage:|exactly one|one folder'
        [IO.File]::Exists($case.Capture) | Should -BeFalse
    }

    It 'uses the actual child exit when an inherited ERRORLEVEL environment variable is present' {
        $case = New-LauncherCase 'install' 'source' 2
        $case.Environment['ERRORLEVEL'] = '99'
        $result = Invoke-LauncherCommand -BatchPath $case.Batch -SourceArguments @($case.Source) -ChildEnvironment $case.Environment
        $result.ExitCode | Should -Be 2
        $result.Stdout | Should -Match '(?i)partial success|master.*(?:retained|safe)'
        [void](Assert-LauncherReceiver $case)
        [Environment]::GetEnvironmentVariable('ERRORLEVEL', 'Process') | Should -BeExactly $parentErrorLevel
    }

    It 'reports a missing adjacent script safely in a metacharacter installation directory' {
        $case = New-LauncherCase 'install ! & (group) [set]' 'source'
        # Delete only the newly created synthetic receiver in this run-owned case.
        [IO.File]::Delete($case.Script)
        $result = Invoke-LauncherCommand -BatchPath $case.Batch -SourceArguments @($case.Source) -ChildEnvironment $case.Environment
        $result.ExitCode | Should -Be 1
        $result.Stdout | Should -Match '(?i)(?:cannot|can.t|could not).*find.*WinPDFMerge\.ps1|WinPDFMerge\.ps1.*missing'
        [IO.File]::Exists($case.Capture) | Should -BeFalse
    }
}
