# Actual Windows entry/dependency preflight faults, plus one separate real PDFtk
# smoke. Controlled version processes prove rejection only, never PDF support.
# Short ASCII/no-space paths keep this tier independent of later native argument,
# publication, email, fidelity and final-page-order acceptance.
param([Parameter(Mandatory=$true)][string]$PdftkPath)

BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
    . (Join-Path $repo 'tests/TestSupport.ps1')
    if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) { throw 'This entry tier requires actual Windows.' }
    if (-not (Test-Path -LiteralPath $PdftkPath -PathType Leaf) -or [IO.Path]::GetFileName($PdftkPath) -ine 'pdftk.exe') {
        throw 'Provide the real vendor pdftk.exe, never the controlled version fixture.'
    }
    $PdftkPath = (Resolve-Path -LiteralPath $PdftkPath).ProviderPath
    $shell = [Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    $parentEnvironment = @{}
    foreach ($name in @('PATH', 'ProgramFiles', 'ProgramFiles(x86)', 'PSModulePath', 'GS_OPTIONS', 'WINPDFMERGER_T07_VERSION_MODE', 'WINPDFMERGER_T07_VERSION_RECEIPT', 'WINPDFMERGER_T07_VERSION_PID_RECEIPT', 'WINPDFMERGER_T07_GS_OPTIONS_RECEIPT')) {
        $parentEnvironment[$name] = [Environment]::GetEnvironmentVariable($name, 'Process')
    }
    $parentPolicies = @(Get-ExecutionPolicy -List | ForEach-Object { $_.Scope.ToString() + '=' + $_.ExecutionPolicy.ToString() }) -join ';'
    $work = Join-Path $repo ('tests/.work/dependency-entry/' + [Guid]::NewGuid().ToString('N'))
    if ($work -match '[^\x21-\x7e]') { throw 'This narrow entry tier requires an ASCII checkout path without spaces until native argument fixes.' }
    [void][IO.Directory]::CreateDirectory($work)
    Write-Host ('Real PDFtk SHA256: ' + (Get-FileHash -LiteralPath $PdftkPath -Algorithm SHA256).Hash)
    Write-Host ('Real PDFtk file version: ' + (Get-Item -LiteralPath $PdftkPath).VersionInfo.FileVersion)

    # Reuse only an existing Windows Framework compiler, with generated files in
    # this unique ignored run. No compiler or runtime installation is attempted.
    $compiler = $null
    foreach ($relative in @('Microsoft.NET/Framework64/v4.0.30319/csc.exe', 'Microsoft.NET/Framework/v4.0.30319/csc.exe')) {
        $candidate = Join-Path $env:WINDIR $relative
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { $compiler = $candidate; break }
    }
    if (-not $compiler) { throw 'No existing Windows C# compiler found; this tier does not install one.' }
    $controlledExecutable = Join-Path $work 'VersionProbeFixture.exe'
    $source = Join-Path $PSScriptRoot 'VersionProbeFixture.cs'
    $compiled = Invoke-TestChildProcess -Executable $compiler -Arguments @('/nologo', '/target:exe', '/platform:anycpu', '/optimize+', ('/out:' + $controlledExecutable), $source)
    [IO.File]::WriteAllText((Join-Path $work 'compiler.txt'), ($compiled.Stdout + $compiled.Stderr), [Text.UTF8Encoding]::new($false))
    if ($compiled.ExitCode -ne 0 -or -not [IO.File]::Exists($controlledExecutable)) { throw ('Controlled version fixture compilation failed: ' + $compiled.Stdout + $compiled.Stderr) }
    $receipt = [ordered]@{
        purpose = 'Controlled version/fault executable only; no PDF engine evidence'
        compiler = $compiler
        compiler_version = (Get-Item -LiteralPath $compiler).VersionInfo.FileVersion
        source_sha256 = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash.ToLowerInvariant()
        executable_sha256 = (Get-FileHash -LiteralPath $controlledExecutable -Algorithm SHA256).Hash.ToLowerInvariant()
        built_at_utc = [DateTime]::UtcNow.ToString('o')
    }
    [IO.File]::WriteAllText((Join-Path $work 'build-info.json'), ($receipt | ConvertTo-Json -Depth 3), [Text.UTF8Encoding]::new($false))

    function New-DependencyEntryApplication {
        $run = Join-Path $work ([Guid]::NewGuid().ToString('N'))
        $app = Join-Path $run 'app'
        $source = Join-Path $run 'source'
        $tools = Join-Path $run 'tools'
        $noCommon = Join-Path $run 'no-common-dependencies'
        foreach ($directory in @((Join-Path $app 'src'), $source, $tools, $noCommon)) { [void][IO.Directory]::CreateDirectory($directory) }
        [IO.File]::Copy((Join-Path $repo 'WinPDFMerge.ps1'), (Join-Path $app 'WinPDFMerge.ps1'))
        [IO.File]::Copy((Join-Path $repo 'src/WinPDFMerge.Helpers.ps1'), (Join-Path $app 'src/WinPDFMerge.Helpers.ps1'))
        [IO.File]::Copy((Join-Path $repo 'tests/fixtures/numbered/2.pdf'), (Join-Path $source 'single.pdf'))
        [pscustomobject]@{
            App = $app
            Source = $source
            Entry = (Join-Path $app 'WinPDFMerge.ps1')
            Tools = $tools
            Receipt = (Join-Path $run 'version-arguments.txt')
            PidReceipt = (Join-Path $run 'version-pid.txt')
            # The explicitly selected shell launches by its full path. No system
            # PATH directories are required for application cmdlets.
            Environment = @{
                PATH = $tools
                ProgramFiles = $noCommon
                'ProgramFiles(x86)' = $noCommon
                WINPDFMERGER_T07_VERSION_MODE = ''
                WINPDFMERGER_T07_VERSION_RECEIPT = ''
                WINPDFMERGER_T07_VERSION_PID_RECEIPT = ''
                WINPDFMERGER_T07_GS_OPTIONS_RECEIPT = ''
            }
        }
    }

    function Get-DependencyEntrySourceSnapshot([string]$Source) {
        foreach ($file in @(Get-ChildItem -LiteralPath $Source -File -Force -Recurse | Sort-Object FullName)) {
            [pscustomobject]@{
                Path = $file.FullName
                Hash = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
                Length = $file.Length
                Modified = $file.LastWriteTimeUtc.Ticks
                Attributes = [int]$file.Attributes
            } | ConvertTo-Json -Compress
        }
    }

    function Assert-DependencyEntryParentUnchanged {
        foreach ($name in $parentEnvironment.Keys) {
            [Environment]::GetEnvironmentVariable($name, 'Process') | Should -BeExactly $parentEnvironment[$name]
        }
        (@(Get-ExecutionPolicy -List | ForEach-Object { $_.Scope.ToString() + '=' + $_.ExecutionPolicy.ToString() }) -join ';') | Should -BeExactly $parentPolicies
    }

    function Set-DependencyTestEnvironmentState([string]$Name, [object]$Value) {
        # PS7 can bind null to String.Empty for the .NET string overload.
        # Remove absence explicitly, while retaining an actual empty string.
        if ($null -eq $Value) {
            Remove-Item -LiteralPath ('Env:\' + $Name) -ErrorAction SilentlyContinue
        } else {
            [Environment]::SetEnvironmentVariable($Name, [string]$Value, 'Process')
        }
    }

    function Invoke-DependencyEntry($Application) {
        Invoke-TestChildProcess -Executable $shell -Arguments @('-NoProfile', '-ExecutionPolicy', 'RemoteSigned', '-File', $Application.Entry, $Application.Source) -ChildEnvironment $Application.Environment -TimeoutMilliseconds 20000
    }

    function Assert-RequiredDependencyEntryFailure($Application, [string]$ExpectedDiagnostic, [bool]$CheckProbeExit = $false, [string]$ExpectedExecutablePath = '', [string[]]$ExpectedDetails = @()) {
        $before = @(Get-DependencyEntrySourceSnapshot $Application.Source) -join "`n"
        $watch = [Diagnostics.Stopwatch]::StartNew()
        $result = Invoke-DependencyEntry $Application
        $watch.Stop()
        if ($CheckProbeExit) {
            # Inspect the exact owned fixture PID immediately at entry return.
            # The product probe's five-second bound is separate from the outer
            # twenty-second harness limit; no other process is touched.
            [IO.File]::Exists($Application.PidReceipt) | Should -BeTrue
            $fixturePid = [int]([IO.File]::ReadAllText($Application.PidReceipt))
            $remainingOwnedProcessCount = @(Get-Process -Id $fixturePid -ErrorAction SilentlyContinue).Count
            $fixturePid | Should -BeGreaterThan 0
            $remainingOwnedProcessCount | Should -Be 0
            $watch.ElapsedMilliseconds | Should -BeLessThan 12000
        }
        $result.ExitCode | Should -Be 1 -Because ($result.Stdout + $result.Stderr)
        $text = $result.Stdout + $result.Stderr
        $text | Should -Match $ExpectedDiagnostic
        if ($ExpectedExecutablePath) { $text | Should -Match ([regex]::Escape($ExpectedExecutablePath)) }
        foreach ($detail in $ExpectedDetails) { $text | Should -Match $detail }
        $text | Should -Match '(?i)install PDFtk Server'
        $text | Should -Not -Match '(?im)^SUCCESS:|PDFtk merge OK|Merge completed successfully'
        @(Get-ChildItem -LiteralPath $Application.App -File -Force -Recurse | Where-Object { $_.Extension -in @('.pdf', '.log') }).Count | Should -Be 0
        (@(Get-DependencyEntrySourceSnapshot $Application.Source) -join "`n") | Should -BeExactly $before
        Assert-DependencyEntryParentUnchanged
    }
}

Describe 'AC015: actual Windows entry rejects absent or unusable PDFtk before outputs' {
    It 'rejects missing PDFtk with installation guidance and no PDF or run log' {
        $application = New-DependencyEntryApplication
        Assert-RequiredDependencyEntryFailure $application '(?i)PDFtk.*(?:not found|missing|unavailable)'
    }

    It 'rejects a discovered pdftk.exe file that is not an executable before outputs' {
        $application = New-DependencyEntryApplication
        $selectedPath = Join-Path $application.Tools 'pdftk.exe'
        [IO.File]::WriteAllText($selectedPath, 'T07 intentionally invalid executable; not a PDF engine.')
        Assert-RequiredDependencyEntryFailure $application '(?i)PDFtk' -ExpectedExecutablePath $selectedPath
    }

    It 'rejects the controlled <Label> version probe before outputs' -TestCases @(
        @{ Label = 'nonzero-exit'; Mode = 'nonzero'; Details = @('(?i)exit(?: code)?\s*7\b', 'T07 controlled version failure\.') },
        @{ Label = 'unrecognized-banner'; Mode = 'unrecognized'; Details = @('T07 unrelated executable, version 2\.02') },
        @{ Label = 'finite timeout'; Mode = 'timeout'; Details = @('(?i)timed out|timeout') }
    ) {
        param($Label, $Mode, $Details)
        $application = New-DependencyEntryApplication
        $selectedPath = Join-Path $application.Tools 'pdftk.exe'
        [IO.File]::Copy($controlledExecutable, $selectedPath)
        $application.Environment['WINPDFMERGER_T07_VERSION_MODE'] = $Mode
        $application.Environment['WINPDFMERGER_T07_VERSION_RECEIPT'] = $application.Receipt
        $application.Environment['WINPDFMERGER_T07_VERSION_PID_RECEIPT'] = $application.PidReceipt
        Assert-RequiredDependencyEntryFailure $application '(?i)PDFtk' -CheckProbeExit ($Mode -eq 'timeout') -ExpectedExecutablePath $selectedPath -ExpectedDetails $Details
        # Exactly one fixed version argument, not a shell or merge invocation.
        [IO.File]::Exists($application.Receipt) | Should -BeTrue
        (@(Get-Content -LiteralPath $application.Receipt) -join "`n") | Should -BeExactly '--version'
    }
}

Describe 'T07 actual entry dependency diagnostics with real PDFtk, narrow native smoke' {
    It 'reports selected real PDFtk path and version and merges two pages with optional GS absent' {
        $application = New-DependencyEntryApplication
        $application.Environment['PATH'] = Split-Path -Parent $PdftkPath
        $before = @(Get-DependencyEntrySourceSnapshot $application.Source) -join "`n"
        $result = Invoke-DependencyEntry $application
        $result.ExitCode | Should -Be 0 -Because ($result.Stdout + $result.Stderr)
        $selectedDiagnostic = '(?m)^PDFtk: ' + [regex]::Escape($PdftkPath) + ' \(version 2\.02\)\r?$'
        $result.Stdout | Should -Match $selectedDiagnostic
        $result.Stdout | Should -Match 'Ghostscript not found; skipping email-optimized copy'
        $logs = @(Get-ChildItem -LiteralPath $application.App -Filter '*.log' -File)
        $logs.Count | Should -Be 1
        $log = Get-Content -LiteralPath $logs[0].FullName -Raw
        $log | Should -Match $selectedDiagnostic
        $masters = @(Get-ChildItem -LiteralPath $application.App -Filter '*.pdf' -File)
        $masters.Count | Should -Be 1
        $masters[0].Name | Should -Not -Match '_email\.pdf$'
        $inspection = Invoke-TestChildProcess -Executable $PdftkPath -Arguments @($masters[0].FullName, 'dump_data_utf8', 'output', '-', 'dont_ask') -TimeoutMilliseconds 10000
        $inspection.ExitCode | Should -Be 0 -Because $inspection.Stderr
        $counts = @([regex]::Matches($inspection.Stdout, '(?m)^NumberOfPages:\s*([0-9]+)\s*$'))
        $counts.Count | Should -Be 1
        [int]$counts[0].Groups[1].Value | Should -Be 2
        (@(Get-DependencyEntrySourceSnapshot $application.Source) -join "`n") | Should -BeExactly $before
        Assert-DependencyEntryParentUnchanged
    }

    It 'retains a real two-page master and returns partial success when a controlled GS version probe fails' {
        # This tests only found-GS version preflight failure. It does not run a
        # real Ghostscript converter or certify later email failure handling.
        $application = New-DependencyEntryApplication
        $selectedGsPath = Join-Path $application.Tools 'gswin64c.exe'
        [IO.File]::Copy($controlledExecutable, $selectedGsPath)
        $application.Environment['PATH'] = (Split-Path -Parent $PdftkPath) + ';' + $application.Tools
        $application.Environment['WINPDFMERGER_T07_VERSION_MODE'] = 'nonzero'
        $application.Environment['WINPDFMERGER_T07_VERSION_RECEIPT'] = $application.Receipt
        $gsOptionsReceipt = Join-Path (Split-Path -Parent $application.Receipt) 'gs-options.txt'
        $application.Environment['WINPDFMERGER_T07_GS_OPTIONS_RECEIPT'] = $gsOptionsReceipt
        $application.Environment['GS_OPTIONS'] = 'T07 child-only version-probe sentinel'
        $before = @(Get-DependencyEntrySourceSnapshot $application.Source) -join "`n"
        $result = Invoke-DependencyEntry $application
        $result.ExitCode | Should -Be 2 -Because ($result.Stdout + $result.Stderr)
        ($result.Stdout + $result.Stderr) | Should -Match '(?i)partial success'
        ($result.Stdout + $result.Stderr) | Should -Not -Match '(?im)^SUCCESS:|Email-optimized PDF created'
        ($result.Stdout + $result.Stderr) | Should -Match ([regex]::Escape($selectedGsPath))
        ($result.Stdout + $result.Stderr) | Should -Match '(?i)exit(?: code)?\s*7\b'
        ($result.Stdout + $result.Stderr) | Should -Match 'T07 controlled version failure\.'
        $masters = @(Get-ChildItem -LiteralPath $application.App -Filter '*.pdf' -File)
        $masters.Count | Should -Be 1
        $masters[0].Name | Should -Not -Match '_email\.pdf$'
        $inspection = Invoke-TestChildProcess -Executable $PdftkPath -Arguments @($masters[0].FullName, 'dump_data_utf8', 'output', '-', 'dont_ask') -TimeoutMilliseconds 10000
        $inspection.ExitCode | Should -Be 0 -Because $inspection.Stderr
        $counts = @([regex]::Matches($inspection.Stdout, '(?m)^NumberOfPages:\s*([0-9]+)\s*$'))
        $counts.Count | Should -Be 1
        [int]$counts[0].Groups[1].Value | Should -Be 2
        $logs = @(Get-ChildItem -LiteralPath $application.App -Filter '*.log' -File)
        $logs.Count | Should -Be 1
        $log = Get-Content -LiteralPath $logs[0].FullName -Raw
        $log | Should -Match '(?i)Ghostscript'
        $log | Should -Match ([regex]::Escape($selectedGsPath))
        $log | Should -Match '(?i)exit(?: code)?\s*7\b'
        $log | Should -Match 'T07 controlled version failure\.'
        [IO.File]::Exists($application.Receipt) | Should -BeTrue
        (@(Get-Content -LiteralPath $application.Receipt) -join "`n") | Should -BeExactly '--version'
        [IO.File]::Exists($gsOptionsReceipt) | Should -BeTrue
        ([IO.File]::ReadAllText($gsOptionsReceipt) -in @('', '<unset>')) | Should -BeTrue
        (@(Get-DependencyEntrySourceSnapshot $application.Source) -join "`n") | Should -BeExactly $before
        Assert-DependencyEntryParentUnchanged
    }
}

Describe 'T07 real dependency version-probe helper with a controlled executable, no PDF or GS support claim' {
    BeforeAll { . (Join-Path $repo 'src/WinPDFMerge.Helpers.ps1') }

    It 'clears GS_OPTIONS only in the version child across requested unset, empty and value caller states' {
        $saved = @{}
        $names = @('GS_OPTIONS', 'WINPDFMERGER_T07_VERSION_MODE', 'WINPDFMERGER_T07_VERSION_RECEIPT', 'WINPDFMERGER_T07_VERSION_PID_RECEIPT', 'WINPDFMERGER_T07_GS_OPTIONS_RECEIPT')
        foreach ($name in $names) { $saved[$name] = [Environment]::GetEnvironmentVariable($name, 'Process') }
        try {
            Set-DependencyTestEnvironmentState 'WINPDFMERGER_T07_VERSION_MODE' 'nonzero'
            Set-DependencyTestEnvironmentState 'WINPDFMERGER_T07_VERSION_PID_RECEIPT' $null
            foreach ($state in @(@{ Label = 'unset'; Value = $null }, @{ Label = 'empty'; Value = '' }, @{ Label = 'value'; Value = 'T07 direct caller sentinel' })) {
                $run = Join-Path $work ([Guid]::NewGuid().ToString('N'))
                [void][IO.Directory]::CreateDirectory($run)
                $argumentsReceipt = Join-Path $run 'direct-arguments.txt'
                $optionsReceipt = Join-Path $run 'direct-gs-options.txt'
                Set-DependencyTestEnvironmentState 'WINPDFMERGER_T07_VERSION_RECEIPT' $argumentsReceipt
                Set-DependencyTestEnvironmentState 'WINPDFMERGER_T07_GS_OPTIONS_RECEIPT' $optionsReceipt
                Set-DependencyTestEnvironmentState 'GS_OPTIONS' $state.Value
                # Windows/.NET may normalize an empty setting to absent. Compare
                # the actual observable caller state, preserving that distinction.
                $before = [Environment]::GetEnvironmentVariable('GS_OPTIONS', 'Process')
                $observed = if ($null -eq $before) { 'unset' } elseif ($before -eq '') { 'empty' } else { 'value' }
                Write-Host ('GS_OPTIONS requested ' + $state.Label + '; observed caller state ' + $observed)
                $result = Invoke-DependencyVersionProbe -Path $controlledExecutable
                [Environment]::GetEnvironmentVariable('GS_OPTIONS', 'Process') | Should -BeExactly $before
                $result.ExitCode | Should -Be 7
                $result.Stdout | Should -Match '^pdftk 2\.02 '
                $result.Stderr | Should -Match 'T07 controlled version failure\.'
                [IO.File]::Exists($argumentsReceipt) | Should -BeTrue
                (@(Get-Content -LiteralPath $argumentsReceipt) -join "`n") | Should -BeExactly '--version'
                [IO.File]::Exists($optionsReceipt) | Should -BeTrue
                ([IO.File]::ReadAllText($optionsReceipt) -in @('', '<unset>')) | Should -BeTrue
            }
        } finally {
            foreach ($name in $names) { Set-DependencyTestEnvironmentState $name $saved[$name] }
        }
        Assert-DependencyEntryParentUnchanged
    }

    It 'uses an injected 200ms version-probe timeout and observes its exact fixture PID exited at return' {
        $saved = @{}
        $names = @('WINPDFMERGER_T07_VERSION_MODE', 'WINPDFMERGER_T07_VERSION_RECEIPT', 'WINPDFMERGER_T07_VERSION_PID_RECEIPT', 'WINPDFMERGER_T07_GS_OPTIONS_RECEIPT')
        foreach ($name in $names) { $saved[$name] = [Environment]::GetEnvironmentVariable($name, 'Process') }
        try {
            $run = Join-Path $work ([Guid]::NewGuid().ToString('N'))
            [void][IO.Directory]::CreateDirectory($run)
            $argumentsReceipt = Join-Path $run 'direct-timeout-arguments.txt'
            $pidReceipt = Join-Path $run 'direct-timeout-pid.txt'
            Set-DependencyTestEnvironmentState 'WINPDFMERGER_T07_VERSION_MODE' 'timeout'
            Set-DependencyTestEnvironmentState 'WINPDFMERGER_T07_VERSION_RECEIPT' $argumentsReceipt
            Set-DependencyTestEnvironmentState 'WINPDFMERGER_T07_VERSION_PID_RECEIPT' $pidReceipt
            Set-DependencyTestEnvironmentState 'WINPDFMERGER_T07_GS_OPTIONS_RECEIPT' $null
            $probeError = $null
            $watch = [Diagnostics.Stopwatch]::StartNew()
            try { $null = Invoke-DependencyVersionProbe -Path $controlledExecutable -TimeoutMilliseconds 200 } catch { $probeError = $_ } finally { $watch.Stop() }
            # The fixture remains finite for regression safety. This observes
            # only the owned version-probe process; it never kills any process.
            [IO.File]::Exists($pidReceipt) | Should -BeTrue
            $fixturePid = [int]([IO.File]::ReadAllText($pidReceipt))
            $remainingOwnedProcessCount = @(Get-Process -Id $fixturePid -ErrorAction SilentlyContinue).Count
            $fixturePid | Should -BeGreaterThan 0
            $remainingOwnedProcessCount | Should -Be 0
            $watch.ElapsedMilliseconds | Should -BeLessThan 5000
            $probeError | Should -Not -BeNullOrEmpty
            $probeError.Exception.Message | Should -Match '(?i)timed out|timeout'
            (@(Get-Content -LiteralPath $argumentsReceipt) -join "`n") | Should -BeExactly '--version'
        } finally {
            foreach ($name in $names) { Set-DependencyTestEnvironmentState $name $saved[$name] }
        }
        Assert-DependencyEntryParentUnchanged
    }
}
