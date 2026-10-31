BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
    $helpers = Join-Path $repo 'src/WinPDFMerge.Helpers.ps1'
    . $helpers
    if (-not ('WinPDFMerge.Tests.RecordingConsoleHost' -as [type])) {
        Add-Type -Path (Join-Path $repo 'tests/support/RecordingConsoleHost.cs')
    }
    function Invoke-RecordedPresentation {
        param([string]$EmailState='skipped', [bool]$MasterPublished=$true, [bool]$Interactive=$true, [switch]$RefuseColors)
        $recordingHost = New-Object WinPDFMerge.Tests.RecordingConsoleHost
        $recordingHost.Recorder.RefuseColors = [bool]$RefuseColors
        $runspace = [Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($recordingHost)
        $runspace.Open()
        $pipeline = [Management.Automation.PowerShell]::Create()
        $pipeline.Runspace = $runspace
        $folder = 'D:\redirected Documents with spaces\' + ('long-copyable-path-' * 12) + '\WinPDFMerger'
        $master = $folder + '\actual-master.pdf'; $email = $folder + '\actual-email.pdf'; $log = $folder + '\actual.log'
        $logWork = Join-Path $repo ('tests/.work/presentation/' + [Guid]::NewGuid().ToString('N'))
        [void][IO.Directory]::CreateDirectory($logWork)
        $diskLog = Join-Path $logWork 'run.log'
        try {
            [void]$pipeline.AddScript(@'
param($Helpers,$Folder,$Master,$Email,$Log,$DiskLog,$EmailState,$MasterPublished,$Interactive)
. $Helpers
$presentation = [pscustomobject]@{Interactive=$Interactive}
Write-PdfConsoleHeader $presentation -Version '1.0.1'
Write-PdfRunStage -Stage 'Input inspection' -Timer ([Diagnostics.Stopwatch]::StartNew()) -LiteralPath $DiskLog -Presentation $presentation
Write-PdfConsoleStage $presentation -Stage SCAN -Message '3 documents / 12 pages' -Role Success
Write-PdfConsoleStage $presentation -Stage PREP -Message 'Synthetic warning' -Role Warning
$outcome = Get-PdfMergeOutcome -MasterPublished $MasterPublished -EmailState $EmailState -MasterPath $Master -EmailPath $Email
Write-PdfConsoleStage $presentation -Stage EMAIL -Message $outcome.EmailMessage -Role $(if($outcome.ExitCode -eq 2){'Warning'}else{'Success'})
Write-PdfConsoleResult $presentation -Outcome $outcome -OutputFolder $Folder -LogPath $Log -Detail $outcome.EmailMessage
('Result: {0}; exit code: {1}' -f $outcome.Summary,$outcome.ExitCode) | Write-RunLog -LiteralPath $DiskLog -Append | Out-Null
'@
            ).AddArgument($helpers).AddArgument($folder).AddArgument($master).AddArgument($email).AddArgument($log).AddArgument($diskLog).AddArgument($EmailState).AddArgument($MasterPublished).AddArgument($Interactive)
            $output = @($pipeline.Invoke())
            @($pipeline.Streams.Error).Count | Should -Be 0 -Because ($pipeline.Streams.Error | Out-String)
            [pscustomobject]@{
                Writes=@($recordingHost.Recorder.Writes); Text=(@($recordingHost.Recorder.Writes | ForEach-Object Text) -join "`n")
                Folder=$folder; Master=$master; Email=$email; Log=$log; DiskLog=[IO.File]::ReadAllText($diskLog); Output=$output
            }
        } finally { $pipeline.Dispose(); $runspace.Dispose() }
    }
}

Describe 'Noninteractive command-line switch detection' {
    It 'recognizes the valid noninteractive switch token <Argument>' -TestCases @(
        @{Argument='-NonI'}, @{Argument='-NonIn'}, @{Argument='-NonInt'},
        @{Argument='-NonInte'}, @{Argument='-NonInter'}, @{Argument='-NonIntera'},
        @{Argument='-NonInterac'}, @{Argument='-NonInteract'}, @{Argument='-NonInteracti'},
        @{Argument='-NonInteractiv'}, @{Argument='-NonInteractive'},
        @{Argument='/nonint'}, @{Argument='-NONINTERACTIVE'},
        @{Argument='"-NonInt"'}, @{Argument='"-NonInteractive"'}, @{Argument='"/nOnI"'}
    ) {
        param($Argument)
        Test-PdfNonInteractiveCommandLine -Argument $Argument | Should -BeTrue
    }
    It 'rejects other arguments or embedded command text: <Argument>' -TestCases @(
        @{Argument='-Non'}, @{Argument='-NonInvalid'}, @{Argument='-NonInteractiveExtra'},
        @{Argument='-NoProfile'}, @{Argument='NonInteractive'}, @{Argument=''},
        @{Argument='Write-Host -NonInt'}, @{Argument='-NonInt with extra tokens'},
        @{Argument='"-NonInt'}, @{Argument='-NonInt"'}, @{Argument='"-Non"'}
    ) {
        param($Argument)
        Test-PdfNonInteractiveCommandLine -Argument $Argument | Should -BeFalse
    }
}

Describe 'Live console renderer on a recording PowerShell host' {
    It 'renders segmented cyan framing, stage column, actual paths and host color roles' {
        $record = Invoke-RecordedPresentation -EmailState published
        $record.Text | Should -Match '\+--\[ WINPDFMERGER 1\.0\.1 \].*\[ LOCAL \]--\+'
        $record.Text | Should -Match 'DOCUMENT OPS  //  MERGE CONSOLE'
        $record.Text | Should -Match '\[CHECK \] Input inspection / [0-9.]+ s'
        $record.Text | Should -Match '\[SCAN  \] 3 documents / 12 pages'
        $record.Text | Should -Match '\+--\[ RESULT: SUCCESS \]'
        foreach ($path in @($record.Folder,$record.Master,$record.Email,$record.Log)) {
            @($record.Writes | Where-Object { $_.Text.EndsWith(': ' + $path) }).Count | Should -Be 1
        }
        @($record.Writes | Where-Object { $_.Text.StartsWith('+--[ WINPDFMERGER') })[0].Foreground | Should -Be Cyan
        @($record.Writes | Where-Object { $_.Text.Contains('DOCUMENT OPS') })[0].Foreground | Should -Be White
        @($record.Writes | Where-Object { $_.Text.Contains('3 documents / 12 pages') })[0].Foreground | Should -Be Green
        @($record.Writes | Where-Object { $_.Text.Contains('Synthetic warning') })[0].Foreground | Should -Be Yellow
        $record.DiskLog | Should -Match 'Stage: Input inspection; elapsed:'
        $record.DiskLog | Should -Not -Match '[\x1b]|\+--|\[CHECK|\[ RESULT:'
    }
    It 'shows PARTIAL in amber and retains the master path after email failure' {
        $record = Invoke-RecordedPresentation -EmailState failed
        $record.Text | Should -Match '\+--\[ RESULT: PARTIAL \]'
        $record.Text | Should -Match ([regex]::Escape($record.Master))
        $record.Text | Should -Not -Match ([regex]::Escape($record.Email))
        @($record.Writes | Where-Object { $_.Text.StartsWith('+--[ RESULT:') })[0].Foreground | Should -Be Yellow
    }
    It 'shows FAILURE in red without claiming unpublished master or email paths' {
        $record = Invoke-RecordedPresentation -MasterPublished $false -EmailState not_started
        $record.Text | Should -Match '\+--\[ RESULT: FAILURE \]'
        $record.Text | Should -Not -Match ([regex]::Escape($record.Master))
        $record.Text | Should -Not -Match ([regex]::Escape($record.Email))
        @($record.Writes | Where-Object { $_.Text.StartsWith('+--[ RESULT:') })[0].Foreground | Should -Be Red
    }
    It 'keeps skipped email paths absent and plain output/log free of frames and escape codes' {
        $record = Invoke-RecordedPresentation -Interactive $false
        $record.Text | Should -Match 'WinPDFMerger 1.0.1'
        $record.Text | Should -Match 'Result: SUCCESS; exit code: 0'
        $record.Text | Should -Match ([regex]::Escape($record.Master))
        $record.Text | Should -Not -Match ([regex]::Escape($record.Email))
        $record.Text | Should -Not -Match '[\x1b]|\+--|DOCUMENT OPS'
        $record.DiskLog | Should -Not -Match '[\x1b]|\+--|DOCUMENT OPS'
        @($record.Writes | Where-Object { $_.Foreground -ne [ConsoleColor]::Black }).Count | Should -Be 0
    }
    It 'finishes rendering with complete paths when the host refuses native colors' {
        $record = Invoke-RecordedPresentation -RefuseColors
        $record.Text | Should -Match 'Result: SUCCESS; exit code: 0'
        $record.Text | Should -Match ([regex]::Escape($record.Master))
    }
}
