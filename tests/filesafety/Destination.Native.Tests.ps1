# Genuine Windows destination/identity integration, synthetic PDFs only.
# Counts are narrow structural oracles, not visual/fidelity acceptance.
param(
    [Parameter(Mandatory=$true)][string]$PdftkPath,
    [Parameter(Mandatory=$true)][string]$GhostscriptPath
)

BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
    if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) { throw 'Destination integration requires actual Windows; unavailable evidence is not skipped.' }
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    try {
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        if ($principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { throw 'Run the destination integration as a standard user, not elevated.' }
        $currentSid = $identity.User
    } finally { $identity.Dispose() }
    . (Join-Path $repo 'src/WinPDFMerge.Helpers.ps1')
    . (Join-Path $repo 'tests/TestSupport.ps1')
    $PdftkPath = (Resolve-Path -LiteralPath $PdftkPath).ProviderPath
    $GhostscriptPath = (Resolve-Path -LiteralPath $GhostscriptPath).ProviderPath
    if ([IO.Path]::GetFileName($PdftkPath) -ine 'pdftk.exe' -or [IO.Path]::GetFileName($GhostscriptPath) -ine 'gswin64c.exe') { throw 'Supply the real approved vendor engines, never controlled fixtures.' }
    $pdftkReceipt = Get-Content -LiteralPath (Join-Path $repo 'docs/codex/evidence/T03-pdftk-acquisition.json') -Raw | ConvertFrom-Json
    $gsReceipt = Get-Content -LiteralPath (Join-Path $repo 'docs/codex/evidence/T09-gs-acquisition.json') -Raw | ConvertFrom-Json
    foreach ($selection in @(
        @{ Path = $PdftkPath; Files = $pdftkReceipt.extracted_files; Leaf = 'pdftk.exe' },
        @{ Path = (Join-Path ([IO.Path]::GetDirectoryName($PdftkPath)) 'libiconv2.dll'); Files = $pdftkReceipt.extracted_files; Leaf = 'libiconv2.dll' },
        @{ Path = $GhostscriptPath; Files = $gsReceipt.ghostscript_extraction.selected_files; Leaf = 'gswin64c.exe' },
        @{ Path = (Join-Path ([IO.Path]::GetDirectoryName($GhostscriptPath)) 'gsdll64.dll'); Files = $gsReceipt.ghostscript_extraction.selected_files; Leaf = 'gsdll64.dll' }
    )) {
        $expected = @($selection.Files | Where-Object relative_path -like ('*/' + $selection.Leaf))
        if ($expected.Count -ne 1 -or -not [IO.File]::Exists($selection.Path) -or
            (Get-FileHash -LiteralPath $selection.Path -Algorithm SHA256).Hash.ToLowerInvariant() -cne $expected[0].sha256) { throw ('Engine or interpreter does not match approved acquisition receipt: ' + $selection.Leaf) }
    }
    $shell = [Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    $work = Join-Path $repo ('tests/.work/destination/' + [Guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($work)
    $observations = New-Object 'System.Collections.Generic.List[object]'
    $fixture = Join-Path $repo 'tests/fixtures/numbered/2.pdf'
    $pdftkVersion = Get-NativeToolVersion -Path $PdftkPath -Tool PdfTk
    $gsVersion = Get-NativeToolVersion -Path $GhostscriptPath -Tool Ghostscript
    $parentPath = [Environment]::GetEnvironmentVariable('PATH', 'Process')
    $parentGsOptions = [Environment]::GetEnvironmentVariable('GS_OPTIONS', 'Process')

    function New-DestinationApplication([string]$Name = 'source') {
        $root = Join-Path $work ([Guid]::NewGuid().ToString('N'))
        $app = Join-Path $root 'app'
        $source = Join-Path $root $Name
        $output = Join-Path $root 'output'
        $noCommon = Join-Path $root 'no-common-engines'
        foreach ($directory in @((Join-Path $app 'src'), $source, $output, $noCommon)) { [void][IO.Directory]::CreateDirectory($directory) }
        [IO.File]::Copy((Join-Path $repo 'WinPDFMerge.ps1'), (Join-Path $app 'WinPDFMerge.ps1'), $false)
        [IO.File]::Copy((Join-Path $repo 'VERSION'), (Join-Path $app 'VERSION'), $false)
        [IO.File]::Copy((Join-Path $repo 'src/WinPDFMerge.Helpers.ps1'), (Join-Path $app 'src/WinPDFMerge.Helpers.ps1'), $false)
        $fixtureInput = Join-Path $source 'input.pdf'
        [IO.File]::Copy($fixture, $fixtureInput, $false)
        [pscustomobject]@{
            Root = $root; App = $app; Source = $source; Output = $output; Input = $fixtureInput
            Entry = (Join-Path $app 'WinPDFMerge.ps1')
            ChildEnvironment = @{ ProgramFiles = $noCommon; 'ProgramFiles(x86)' = $noCommon; GS_OPTIONS = '-T10-invalid-inherited-child-option' }
        }
    }

    function Get-DestinationSnapshot([string[]]$Paths) {
        (@($Paths | ForEach-Object {
            $file = Get-Item -LiteralPath $_ -Force
            [pscustomobject]@{ Path = $file.FullName; Hash = (Get-FileHash -LiteralPath $_ -Algorithm SHA256).Hash; Length = $file.Length; Modified = $file.LastWriteTimeUtc.Ticks; Attributes = [int]$file.Attributes } | ConvertTo-Json -Compress
        }) -join "`n")
    }

    function Get-DestinationChildPath([bool]$WithTools = $true) {
        $path = Join-Path $env:SystemRoot 'System32'
        if ($WithTools) { $path = [IO.Path]::GetDirectoryName($PdftkPath) + ';' + [IO.Path]::GetDirectoryName($GhostscriptPath) + ';' + $path }
        $path
    }

    function Invoke-DestinationEntry($Application, [string]$SourcePath, [string]$OutputPath, [bool]$WithTools = $true) {
        if (-not $SourcePath) { $SourcePath = $Application.Source }
        $arguments = @('-NoProfile', '-ExecutionPolicy', 'RemoteSigned', '-File', $Application.Entry, $SourcePath)
        if ($PSBoundParameters.ContainsKey('OutputPath')) { $arguments += @('-OutputFolder', $OutputPath) }
        Invoke-TestChildProcess -Executable $shell -Arguments $arguments -ChildPath (Get-DestinationChildPath $WithTools) -ChildEnvironment $Application.ChildEnvironment -TimeoutMilliseconds 30000
    }

    function Add-DestinationObservation([string]$Label, $Result, $Application, [string]$OutputPath, [string]$Snapshot, [string]$Log = '') {
        $processIdProperty = $Result.PSObject.Properties['ProcessId']
        $observations.Add([pscustomobject]@{
            Label = $Label; ExitCode = $Result.ExitCode; Stdout = $Result.Stdout; Stderr = $Result.Stderr
            Source = $Application.Source; Output = $OutputPath; SourceSnapshot = $Snapshot; Log = $Log
            ProcessId = if ($null -eq $processIdProperty) { $null } else { $processIdProperty.Value }
        })
    }

    function Assert-NoDestinationResidue([string]$Directory) {
        @(Get-ChildItem -LiteralPath $Directory -Force | Where-Object { $_.Name -like '.WinPDFMerge*' }).Count | Should -Be 0
    }

    function Assert-DestinationPageCount([string]$Path) {
        $result = Invoke-TestChildProcess -Executable $PdftkPath -Arguments @($Path, 'dump_data_utf8', 'output', '-', 'dont_ask') -TimeoutMilliseconds 10000
        $result.ExitCode | Should -Be 0 -Because $result.Stderr
        $counts = @([regex]::Matches($result.Stdout, '(?m)^NumberOfPages:\s*([0-9]+)\s*$'))
        $counts.Count | Should -Be 1
        [int]$counts[0].Groups[1].Value | Should -Be 2
    }

    function Assert-DestinationSuccess($Application, $Result, [string]$Directory, [string[]]$ExistingPaths = @()) {
        $Result.ExitCode | Should -Be 0 -Because ($Result.Stdout + $Result.Stderr)
        $newFiles = @(Get-ChildItem -LiteralPath $Directory -File | Where-Object FullName -notin $ExistingPaths)
        $masters = @($newFiles | Where-Object { $_.Name -like 'WinPDFMerge_*.pdf' -and $_.Name -notlike '*_email.pdf' })
        $masters.Count | Should -Be 1
        $stem = $masters[0].BaseName
        $stem | Should -Match '^WinPDFMerge_.+_[0-9]{8}_[0-9]{6}_[0-9a-fA-F]{16}$'
        $email = Join-Path $Directory ($stem + '_email.pdf')
        $logPath = Join-Path $Directory ($stem + '.log')
        [IO.File]::Exists($logPath) | Should -BeTrue
        Assert-DestinationPageCount $masters[0].FullName
        $log = [IO.File]::ReadAllText($logPath, [Text.Encoding]::UTF8)
        if ([IO.File]::Exists($email)) {
            Assert-DestinationPageCount $email
            (Get-Item -LiteralPath $email).Length | Should -BeLessThan $masters[0].Length
        } else {
            $log | Should -Match '(?i)no size benefit'
            $Result.Stdout | Should -Not -Match '(?m)^ - Email'
        }
        $log | Should -Match '(?m)^PDFtk stdout:'
        $log | Should -Match '(?m)^Ghostscript stdout:'
        $log | Should -Match ([regex]::Escape($masters[0].FullName))
        Assert-NoDestinationResidue $Directory
        $log
    }

    function Start-DestinationEntry($Application, [string]$Directory) {
        $arguments = @('-NoProfile', '-ExecutionPolicy', 'RemoteSigned', '-File', $Application.Entry, $Application.Source, '-OutputFolder', $Directory)
        $rendered = foreach ($argument in $arguments) {
            if ($argument.Contains('"')) { throw 'Synthetic Windows paths must not contain quotes.' }
            '"' + ($argument -replace '(\\+)$', '$1$1') + '"'
        }
        $info = New-Object Diagnostics.ProcessStartInfo
        $info.FileName = $shell
        $info.Arguments = $rendered -join ' '
        $info.UseShellExecute = $false
        $info.CreateNoWindow = $true
        $info.RedirectStandardInput = $true
        $info.RedirectStandardOutput = $true
        $info.RedirectStandardError = $true
        $info.StandardOutputEncoding = [Text.Encoding]::UTF8
        $info.StandardErrorEncoding = [Text.Encoding]::UTF8
        $info.EnvironmentVariables['PATH'] = Get-DestinationChildPath
        foreach ($name in $Application.ChildEnvironment.Keys) { $info.EnvironmentVariables[$name] = $Application.ChildEnvironment[$name] }
        $process = New-Object Diagnostics.Process
        $process.StartInfo = $info
        if (-not $process.Start()) { throw 'Actual concurrent application child did not start.' }
        $process.StandardInput.Close()
        [pscustomobject]@{ Process = $process; Stdout = $process.StandardOutput.ReadToEndAsync(); Stderr = $process.StandardError.ReadToEndAsync() }
    }

    function Complete-DestinationEntry($Child) {
        if (-not $Child.Process.WaitForExit(30000)) {
            $Child.Process.Kill()
            if (-not $Child.Process.WaitForExit(5000)) { throw 'Exact owned application child could not be stopped.' }
            throw 'Concurrent application child exceeded its finite limit.'
        }
        if (-not [Threading.Tasks.Task]::WaitAll([Threading.Tasks.Task[]]@($Child.Stdout, $Child.Stderr), 5000)) {
            throw 'Concurrent application stream capture exceeded its finite limit after process exit; inherited pipes may remain open.'
        }
        [pscustomobject]@{ ExitCode = $Child.Process.ExitCode; Stdout = $Child.Stdout.Result; Stderr = $Child.Stderr.Result; ProcessId = $Child.Process.Id }
    }

    function Remove-DestinationJunction([string]$Link) {
        $item = Get-Item -LiteralPath $Link -Force
        if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -eq 0 -or
            -not $item.FullName.StartsWith($work + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Refusing cleanup of anything except this suite-owned junction.' }
        # RemoveDirectory on the reparse point unlinks only that junction; no
        # recursive delete or traversal of the target's files is performed.
        [IO.Directory]::Delete($item.FullName, $false)
    }

    function Get-DestinationShortPath([string]$Path) {
        if (-not ('T10DirectoryAliasProbe' -as [type])) {
            Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class T10DirectoryAliasProbe {
    [DllImport("kernel32.dll", EntryPoint="GetShortPathNameW", CharSet=CharSet.Unicode, SetLastError=true)]
    public static extern uint GetShortPathName(string path, StringBuilder result, uint size);
}
'@
        }
        $buffer = New-Object Text.StringBuilder(1024)
        $length = [T10DirectoryAliasProbe]::GetShortPathName($Path, $buffer, [uint32]$buffer.Capacity)
        if ($length -eq 0 -or $length -ge $buffer.Capacity) { return $null }
        $buffer.ToString()
    }
}

AfterAll {
    [Environment]::GetEnvironmentVariable('PATH', 'Process') | Should -BeExactly $parentPath
    [Environment]::GetEnvironmentVariable('GS_OPTIONS', 'Process') | Should -BeExactly $parentGsOptions
    $report = Join-Path $work 'native-observations.json'
    [ordered]@{
        CommitUnderTest = (& git -C $repo rev-parse HEAD); DirtyWorktree = (@(& git -C $repo status --porcelain=v1).Count -ne 0)
        ShellVersion = $PSVersionTable.PSVersion.ToString(); ShellEdition = $PSVersionTable.PSEdition
        StandardUser = $true; PdfTkVersion = $pdftkVersion; GhostscriptVersion = $gsVersion
        Observations = $observations.ToArray(); Scope = 'Owned synthetic local NTFS destination/identity/ACL/concurrency evidence; no desktop or fidelity claim.'
    } | ConvertTo-Json -Depth 8 | Write-RunLog -LiteralPath $report | Out-Null
    Write-Host ('Native observations: ' + $report)
}

Describe 'AC022: actual entry destination preflight and standard-user writability' {
    It 'keeps the omitted destination beside the entry script and uses actual engines without editing sources' {
        $application = New-DestinationApplication
        $before = Get-DestinationSnapshot @($application.Input)
        $result = Invoke-DestinationEntry $application
        $log = Assert-DestinationSuccess $application $result $application.App
        @(Get-ChildItem -LiteralPath $application.Output -Force).Count | Should -Be 0
        (Get-DestinationSnapshot @($application.Input)) | Should -BeExactly $before
        Add-DestinationObservation 'default-entry-directory-writable' $result $application $application.App $before $log
    }

    It 'accepts a named literal writable OutputFolder with punctuation and Latin Unicode' {
        $application = New-DestinationApplication
        $output = Join-Path $application.Root ("output [x] & ! (a) apostrophe's " + [char]0x00e4)
        [void][IO.Directory]::CreateDirectory($output)
        $foreign = Join-Path $output 'foreign-owned.txt'
        [IO.File]::WriteAllText($foreign, 'T10 synthetic foreign output sentinel')
        $before = Get-DestinationSnapshot @($application.Input, $foreign)
        $result = Invoke-DestinationEntry $application -OutputPath $output
        $log = Assert-DestinationSuccess $application $result $output @($foreign)
        @(Get-ChildItem -LiteralPath $application.App -File -Filter 'WinPDFMerge_*').Count | Should -Be 0
        (Get-DestinationSnapshot @($application.Input, $foreign)) | Should -BeExactly $before
        Add-DestinationObservation 'explicit-literal-writable-output' $result $application $output $before $log
    }

    It 'rejects a <Kind> destination before dependency lookup, names or outputs' -TestCases @(
        @{ Kind = 'nonexistent' }, @{ Kind = 'file' }, @{ Kind = 'wildcard' }, @{ Kind = 'nonfilesystem provider' }
    ) {
        param($Kind)
        $application = New-DestinationApplication
        $paths = @($application.Input)
        switch ($Kind) {
            'nonexistent' { $output = Join-Path $application.Root 'missing-output' }
            'file' { $output = Join-Path $application.Root 'destination-file.txt'; [IO.File]::WriteAllText($output, 'T10 file destination sentinel'); $paths += $output }
            'wildcard' { $output = $application.Output + '*' }
            'nonfilesystem provider' { $output = 'Env:\' }
        }
        $before = Get-DestinationSnapshot $paths
        $result = Invoke-DestinationEntry $application -OutputPath $output -WithTools $false
        $result.ExitCode | Should -Be 1
        ($result.Stdout + $result.Stderr) | Should -Match '(?i)output|destination|FileSystem|wildcard'
        ($result.Stdout + $result.Stderr) | Should -Not -Match 'PDFtk preflight failed'
        @(Get-ChildItem -LiteralPath $application.App -File -Filter 'WinPDFMerge_*').Count | Should -Be 0
        Assert-NoDestinationResidue $application.Output
        if ($Kind -eq 'nonexistent') { [IO.Directory]::Exists($output) | Should -BeFalse }
        (Get-DestinationSnapshot $paths) | Should -BeExactly $before
        Add-DestinationObservation ('invalid-output-' + $Kind) $result $application $output $before
    }

    It 'refuses a denied-write default installation and recovers through an explicit writable destination' {
        $application = New-DestinationApplication
        $originalAcl = Get-Acl -LiteralPath $application.App
        $originalDescriptor = $originalAcl.Sddl
        $deniedAcl = Get-Acl -LiteralPath $application.App
        $deny = New-Object Security.AccessControl.FileSystemAccessRule($currentSid, [Security.AccessControl.FileSystemRights]::WriteData, [Security.AccessControl.AccessControlType]::Deny)
        [void]$deniedAcl.AddAccessRule($deny)
        $before = Get-DestinationSnapshot @($application.Input, $application.Entry)
        try {
            Set-Acl -LiteralPath $application.App -AclObject $deniedAcl
            [IO.File]::ReadAllText($application.Entry).Length | Should -BeGreaterThan 0
            $probe = Join-Path $application.App 'T10-owned-denial-proof.tmp'
            { $file = [IO.File]::Open($probe, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None); $file.Dispose() } | Should -Throw
            $failure = Invoke-DestinationEntry $application -WithTools $false
            $failure.ExitCode | Should -Be 1
            ($failure.Stdout + $failure.Stderr) | Should -Match '(?i)OutputFolder|writ|output|destination'
            ($failure.Stdout + $failure.Stderr) | Should -Not -Match 'PDFtk preflight failed'
            @(Get-ChildItem -LiteralPath $application.App -File -Filter 'WinPDFMerge_*').Count | Should -Be 0
            Assert-NoDestinationResidue $application.App
            $success = Invoke-DestinationEntry $application -OutputPath $application.Output
            $log = Assert-DestinationSuccess $application $success $application.Output
            Add-DestinationObservation 'default-actual-directory-ACL-denial' $failure $application $application.App $before
            Add-DestinationObservation 'explicit-writable-recovery-from-restricted-install' $success $application $application.Output $before $log
        } finally { Set-Acl -LiteralPath $application.App -AclObject $originalAcl }
        (Get-Acl -LiteralPath $application.App).Sddl | Should -BeExactly $originalDescriptor
        (Get-DestinationSnapshot @($application.Input, $application.Entry)) | Should -BeExactly $before
    }

    It 'removes its successful owned writability probe when later dependency preflight fails' {
        $application = New-DestinationApplication
        $foreign = Join-Path $application.Output 'foreign.tmp'
        [IO.File]::WriteAllText($foreign, 'T10 unrelated temp file must survive')
        $before = Get-DestinationSnapshot @($application.Input, $foreign)
        $result = Invoke-DestinationEntry $application -OutputPath $application.Output -WithTools $false
        $result.ExitCode | Should -Be 1
        ($result.Stdout + $result.Stderr) | Should -Match 'PDFtk preflight failed'
        Assert-NoDestinationResidue $application.Output
        @(Get-ChildItem -LiteralPath $application.Output -File).Count | Should -Be 2
        $logs = @(Get-ChildItem -LiteralPath $application.Output -File -Filter 'WinPDFMerge_*.log')
        $logs.Count | Should -Be 1
        $log = [IO.File]::ReadAllText($logs[0].FullName)
        $log | Should -Match 'PDFtk preflight failed'
        $log | Should -Match 'Stage: PDFtk preflight; elapsed:'
        $log | Should -Match 'Result: Failure; exit code: 1'
        $log | Should -Not -Match 'PDFtk version probe executable:|Published Merged master:'
        (Get-DestinationSnapshot @($application.Input, $foreign)) | Should -BeExactly $before
        Add-DestinationObservation 'owned-probe-cleanup-before-later-dependency-failure' $result $application $application.Output $before
    }
}

Describe 'AC023: real directory identity and NTFS junction protection' {
    It 'refuses source/output <Kind> before any native dependency or output creation' -TestCases @(
        @{ Kind = 'same directory' }, @{ Kind = 'case alias' }
    ) {
        param($Kind)
        $application = New-DestinationApplication
        $before = Get-DestinationSnapshot @($application.Input)
        $output = if ($Kind -eq 'case alias') { $application.Source.ToUpperInvariant() } else { $application.Source }
        $result = Invoke-DestinationEntry $application -OutputPath $output -WithTools $false
        $result.ExitCode | Should -Be 1
        ($result.Stdout + $result.Stderr) | Should -Match '(?i)same|overlap|separate|identity'
        ($result.Stdout + $result.Stderr) | Should -Not -Match 'PDFtk preflight failed'
        @(Get-ChildItem -LiteralPath $application.Source -File).Count | Should -Be 1
        Assert-NoDestinationResidue $application.Source
        (Get-DestinationSnapshot @($application.Input)) | Should -BeExactly $before
        Add-DestinationObservation ('directory-overlap-' + $Kind) $result $application $output $before
        if ($Kind -eq 'same directory') {
            $shortPath = Get-DestinationShortPath $application.Source
            if ($shortPath -and $shortPath -ine $application.Source -and [IO.Directory]::Exists($shortPath)) {
                $shortResult = Invoke-DestinationEntry $application -OutputPath $shortPath -WithTools $false
                $shortResult.ExitCode | Should -Be 1
                ($shortResult.Stdout + $shortResult.Stderr) | Should -Match '(?i)same|overlap|separate|identity'
                ($shortResult.Stdout + $shortResult.Stderr) | Should -Not -Match 'PDFtk preflight failed'
                (Get-DestinationSnapshot @($application.Input)) | Should -BeExactly $before
                Add-DestinationObservation 'available-actual-8.3-directory-alias-refused' $shortResult $application $shortPath $before
            } else {
                $observations.Add([pscustomobject]@{ Label = 'optional-8.3-directory-alias'; Result = 'not_run'; Reason = 'GetShortPathNameW did not return a distinct existing alias; volume policy was not changed.'; ObservedPath = $shortPath })
            }
        }
    }

    It 'refuses an actual <Kind> junction without traversing, self-merging or touching its target' -TestCases @(
        @{ Kind = 'output leaf to source' }, @{ Kind = 'source leaf to output' },
        @{ Kind = 'output ancestor' }, @{ Kind = 'source ancestor' }
    ) {
        param($Kind)
        $application = New-DestinationApplication
        $source = $application.Source
        $output = $application.Output
        $link = Join-Path $application.Root 'owned-junction'
        $paths = @($application.Input)
        switch ($Kind) {
            'output leaf to source' { $target = $application.Source; $output = $link }
            'source leaf to output' {
                $target = $application.Output
                $fixtureInput = Join-Path $target 'input.pdf'
                [IO.File]::Copy($fixture, $fixtureInput, $false)
                $paths += $fixtureInput
                $source = $link
            }
            'output ancestor' {
                $target = Join-Path $application.Root 'real-output-parent'
                [void][IO.Directory]::CreateDirectory((Join-Path $target 'nested'))
                $output = Join-Path $link 'nested'
            }
            'source ancestor' {
                $target = Join-Path $application.Root 'real-source-parent'
                [void][IO.Directory]::CreateDirectory((Join-Path $target 'nested'))
                $fixtureInput = Join-Path $target 'nested/input.pdf'
                [IO.File]::Copy($fixture, $fixtureInput, $false)
                $paths += $fixtureInput
                $source = Join-Path $link 'nested'
            }
        }
        $null = New-Item -ItemType Junction -Path $link -Target $target -ErrorAction Stop
        ((Get-Item -LiteralPath $link -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) | Should -Be ([IO.FileAttributes]::ReparsePoint)
        $before = Get-DestinationSnapshot $paths
        try {
            $result = Invoke-DestinationEntry $application -SourcePath $source -OutputPath $output -WithTools $false
            $result.ExitCode | Should -Be 1
            ($result.Stdout + $result.Stderr) | Should -Match '(?i)reparse|junction|unsupported|alias'
            ($result.Stdout + $result.Stderr) | Should -Not -Match 'PDFtk preflight failed'
            @(Get-ChildItem -LiteralPath $application.App -File -Filter 'WinPDFMerge_*').Count | Should -Be 0
            Assert-NoDestinationResidue $target
            (Get-DestinationSnapshot $paths) | Should -BeExactly $before
            Add-DestinationObservation ('actual-junction-' + $Kind) $result $application $output $before
        } finally { Remove-DestinationJunction $link }
        [IO.Directory]::Exists($target) | Should -BeTrue
        (Get-DestinationSnapshot $paths) | Should -BeExactly $before
    }
}

Describe 'AC024: actual simultaneous run identities and unchanged existing results' {
    It 'gives overlapping real application runs distinct shared master/email/log identities in one output directory' {
        $application = New-DestinationApplication
        $oldStem = 'WinPDFMerge_source_20000101_000000_deadbeefdeadbeef'
        $existing = @()
        foreach ($suffix in @('.pdf', '_email.pdf', '.log')) {
            $path = Join-Path $application.Output ($oldStem + $suffix)
            [IO.File]::WriteAllText($path, 'T10 synthetic existing output bytes: ' + $suffix)
            $existing += $path
        }
        $before = Get-DestinationSnapshot (@($application.Input) + $existing)
        $first = $null
        $second = $null
        try {
            $first = Start-DestinationEntry $application $application.Output
            $second = Start-DestinationEntry $application $application.Output
            $first.Process.HasExited | Should -BeFalse
            $second.Process.HasExited | Should -BeFalse
            $one = Complete-DestinationEntry $first
            $two = Complete-DestinationEntry $second
            $one.ExitCode | Should -Be 0 -Because ($one.Stdout + $one.Stderr)
            $two.ExitCode | Should -Be 0 -Because ($two.Stdout + $two.Stderr)
            $newFiles = @(Get-ChildItem -LiteralPath $application.Output -File | Where-Object FullName -notin $existing)
            $masters = @($newFiles | Where-Object { $_.Name -like '*.pdf' -and $_.Name -notlike '*_email.pdf' })
            $masters.Count | Should -Be 2
            @($masters | Select-Object -ExpandProperty BaseName -Unique).Count | Should -Be 2
            $emailFiles=@($newFiles | Where-Object Name -like '*_email.pdf')
            $newFiles.Count | Should -Be (4 + $emailFiles.Count)
            foreach ($master in $masters) {
                $master.BaseName | Should -Match '^WinPDFMerge_.+_[0-9]{8}_[0-9]{6}_[0-9a-fA-F]{16}$'
                $email = Join-Path $application.Output ($master.BaseName + '_email.pdf')
                $logPath = Join-Path $application.Output ($master.BaseName + '.log')
                [IO.File]::Exists($logPath) | Should -BeTrue
                Assert-DestinationPageCount $master.FullName
                $log = [IO.File]::ReadAllText($logPath, [Text.Encoding]::UTF8)
                if ([IO.File]::Exists($email)) {
                    Assert-DestinationPageCount $email
                    (Get-Item -LiteralPath $email).Length | Should -BeLessThan $master.Length
                } else { $log | Should -Match '(?i)no size benefit' }
                $log | Should -Match ([regex]::Escape($master.FullName))
                $log | Should -Match ([regex]::Escape($email))
                $matchingResults = @(@($one, $two) | Where-Object { $_.Stdout.Contains($master.FullName) })
                $matchingResults.Count | Should -Be 1
                foreach ($otherMaster in @($masters | Where-Object FullName -ne $master.FullName)) {
                    $matchingResults[0].Stdout | Should -Not -Match ([regex]::Escape($otherMaster.BaseName))
                    $log | Should -Not -Match ([regex]::Escape($otherMaster.BaseName))
                }
                Add-DestinationObservation ('simultaneous-shared-identity-' + $master.BaseName) $matchingResults[0] $application $application.Output $before $log
            }
            Assert-NoDestinationResidue $application.Output
            (Get-DestinationSnapshot (@($application.Input) + $existing)) | Should -BeExactly $before
            $observations.Add([pscustomobject]@{ Label = 'actual-concurrent-processes'; ProcessIds = @($one.ProcessId, $two.ProcessId); Results = @($one, $two); Identities = @($masters.BaseName); SourcesAndExistingSnapshot = $before })
        } finally {
            foreach ($child in @($first, $second)) {
                if ($null -eq $child) { continue }
                try { if (-not $child.Process.HasExited) { $child.Process.Kill(); [void]$child.Process.WaitForExit(5000) } }
                finally { $child.Process.Dispose() }
            }
        }
    }
}
