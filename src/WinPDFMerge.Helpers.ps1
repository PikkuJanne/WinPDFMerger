# Internal helpers. Import defines functions only; entry orchestration stays in WinPDFMerge.ps1.
# Unchanged baseline helpers retain their measured behavior until their regression task.

function Resolve-SourceDirectory {
    [CmdletBinding()]
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw 'SourceFolder must name exactly one existing FileSystem directory.'
    }
    # Brackets are valid literal filename characters. Only the Win32-invalid
    # wildcard characters are rejected; no wildcard expansion is performed.
    if ($Path.IndexOfAny([char[]]'*?') -ge 0) {
        throw "SourceFolder does not support wildcard expansion: '$Path'. Supply one literal directory."
    }
    try {
        $resolved = @(Resolve-Path -LiteralPath $Path -ErrorAction Stop)
    } catch {
        throw "SourceFolder does not exist or cannot be accessed: '$Path'. $($_.Exception.Message)"
    }
    if ($resolved.Count -ne 1 -or $resolved[0].Provider.Name -ne 'FileSystem') {
        throw "SourceFolder must resolve to exactly one FileSystem directory: '$Path'."
    }
    try {
        # Force permits an explicitly selected hidden directory, not hidden PDFs.
        $directory = Get-Item -LiteralPath $resolved[0].ProviderPath -Force -ErrorAction Stop
    } catch {
        throw "SourceFolder directory cannot be accessed: '$Path'. $($_.Exception.Message)"
    }
    if (-not $directory.PSIsContainer) {
        throw "SourceFolder is not a directory: '$Path'."
    }
    $fullPath = [IO.Path]::GetFullPath($directory.FullName)
    $root = [IO.Path]::GetPathRoot($fullPath)
    if ($fullPath.Length -gt $root.Length) {
        $fullPath = $fullPath.TrimEnd([char[]]'\/')
    }
    return $fullPath
}

function Get-SourcePdfFiles {
    [CmdletBinding()]
    param([string]$SourceFolder)

    $directory = Resolve-SourceDirectory -Path $SourceFolder
    try {
        # Preserve the non-Force, top-level-only scan. Extension comparison is
        # explicitly case-insensitive; directories and wildcard near-matches
        # cannot enter the frozen collection.
        $files = @(Get-ChildItem -LiteralPath $directory -Filter '*.pdf' -File -ErrorAction Stop |
            Where-Object { $_.Extension -ieq '.pdf' })
    } catch {
        throw "Cannot read top-level PDFs from SourceFolder '$directory'. $($_.Exception.Message)"
    }
    if ($files.Count -eq 0) {
        throw "No PDFs found in: '$directory'. Only visible top-level .pdf files are included."
    }
    # Callers wrap this FileInfo stream in @() for the single-input case too.
    return $files
}

function Write-RunLog {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline=$true)][string]$Message,
        [Parameter(Mandatory=$true)][string]$LiteralPath,
        [switch]$Append
    )
    begin { $writeAppend = $Append.IsPresent }
    process {
        # Out-File utf8 has different BOM defaults in PS5.1 and PS7. Write the
        # same UTF-8-without-BOM bytes in both shells, retaining literal paths.
        $encoding = New-Object Text.UTF8Encoding($false)
        $line = $Message + [Environment]::NewLine
        if ($writeAppend) {
            [IO.File]::AppendAllText($LiteralPath, $line, $encoding)
        } else {
            [IO.File]::WriteAllText($LiteralPath, $line, $encoding)
        }
        $writeAppend = $true
        $Message
    }
}

function ConvertTo-NativeArgumentString {
    [CmdletBinding()]
    param(
        # string[] coercion changes null elements into empty strings before
        # validation. Retain object[] here, then require actual strings below.
        [Parameter(Mandatory=$true)][AllowNull()][AllowEmptyCollection()][object[]]$Arguments
    )

    if ($null -eq $Arguments) { throw 'Native arguments must be a string vector; use @() for no arguments.' }
    $rendered = New-Object Text.StringBuilder
    foreach ($argument in $Arguments) {
        if ($null -eq $argument -or $argument -isnot [string]) {
            throw 'Every native argument must be a non-null string; an empty string is allowed.'
        }
        if ($argument.IndexOf([char]0) -ge 0) { throw 'Native arguments cannot contain NUL characters.' }
        if ($rendered.Length -gt 0) { $null = $rendered.Append(' ') }
        # Windows CRT parsing: quote every operand, double backslashes before
        # an embedded quote and before the closing quote, and escape quotes.
        $null = $rendered.Append([char]34)
        $slashes = 0
        foreach ($character in $argument.ToCharArray()) {
            if ($character -eq [char]92) { $slashes++; continue }
            if ($character -eq [char]34) {
                $null = $rendered.Append([char]92, (2 * $slashes + 1))
            } elseif ($slashes -gt 0) {
                $null = $rendered.Append([char]92, $slashes)
            }
            $null = $rendered.Append($character)
            $slashes = 0
        }
        if ($slashes -gt 0) { $null = $rendered.Append([char]92, (2 * $slashes)) }
        $null = $rendered.Append([char]34)
    }
    return $rendered.ToString()
}

function ConvertTo-NativeLogText {
    param([AllowNull()][string]$Text)

    if ($null -eq $Text) { return '' }
    $sanitized = New-Object Text.StringBuilder
    foreach ($character in $Text.ToCharArray()) {
        if ([char]::IsControl($character)) {
            $null = $sanitized.Append(('\u{0:X4}' -f [int]$character))
        } else {
            $null = $sanitized.Append($character)
        }
    }
    return $sanitized.ToString()
}

function Assert-NativeCommandLength {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$Executable,
        [AllowEmptyString()][string]$SerializedArguments = '',
        [ValidateRange(1, 32766)][int]$MaximumCommandLineCharacters = 30000
    )

    # Count UTF-16 code units, executable quotes, separating space and final NUL.
    # Keep headroom beneath CreateProcessW's 32767-character limit.
    $executableText = ConvertTo-NativeArgumentString -Arguments @($Executable)
    $length = $executableText.Length + 1
    if ($SerializedArguments.Length -gt 0) { $length += 1 + $SerializedArguments.Length }
    if ($length -gt $MaximumCommandLineCharacters) {
        throw "Native command requires $length UTF-16 characters including executable, quoting and terminator; the limit is $MaximumCommandLineCharacters. Use fewer inputs or shorter folder paths. No native process was launched; source files were not renamed."
    }
    return $length
}

function New-NativeStreamCapture {
    param([IO.StreamReader]$Reader)

    $state = [pscustomobject]@{
        Reader = $Reader
        Buffer = New-Object 'char[]' 8192
        PendingRead = $null
        Text = New-Object Text.StringBuilder
        Closed = $false
        Truncated = $false
        Error = $null
    }
    $state.PendingRead = $Reader.ReadAsync($state.Buffer, 0, $state.Buffer.Length)
    return $state
}

function Receive-NativeStreamCapture {
    param($State, [int]$MaximumCaptureCharacters)

    if ($null -eq $State -or $State.Closed -or -not $State.PendingRead.IsCompleted) { return $false }
    try {
        # Result is read only after IsCompleted; this never waits for a pipe.
        $length = $State.PendingRead.Result
        if ($length -eq 0) {
            $State.Closed = $true
            $State.PendingRead = $null
            return $true
        }
        $remaining = $MaximumCaptureCharacters - $State.Text.Length
        $retained = [Math]::Min($remaining, $length)
        if ($retained -gt 0) { $null = $State.Text.Append($State.Buffer, 0, $retained) }
        if ($retained -lt $length) { $State.Truncated = $true }
        # Continue draining after the retention limit to avoid blocking a child
        # on full stdout/stderr pipes. One chunk per stream keeps polling fair.
        $State.PendingRead = $State.Reader.ReadAsync($State.Buffer, 0, $State.Buffer.Length)
    } catch {
        $State.Error = $_.Exception.Message
        $State.Closed = $true
        $State.PendingRead = $null
    }
    return $true
}

function Stop-OwnedNativeProcess {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][Diagnostics.Process]$Process,
        [ValidateRange(1, 60000)][int]$TimeoutMilliseconds = 1000
    )

    try {
        if ($Process.HasExited) { return $null }
        # Kill exactly this Process instance. Descendant ownership/cancellation
        # is a later task; never use image-name-wide termination here.
        $Process.Kill()
        if (-not $Process.WaitForExit($TimeoutMilliseconds)) {
            return "Owned native process termination timed out after $TimeoutMilliseconds ms; cleanup was best effort."
        }
    } catch {
        return ('Owned native process termination failed; cleanup was best effort. ' + $_.Exception.Message)
    }
    return $null
}

function Invoke-NativeProcess {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$Executable,
        [AllowNull()][AllowEmptyCollection()][object[]]$Arguments = @(),
        [ValidateRange(1, 2147483647)][int]$TimeoutMilliseconds = 900000,
        [ValidateRange(1, 60000)][int]$TerminationTimeoutMilliseconds = 1000,
        [ValidateRange(1, 60000)][int]$CaptureTimeoutMilliseconds = 1000,
        [ValidateRange(1, 2147483647)][int]$MaximumCaptureCharacters = 8388608,
        [ValidateRange(1, 32766)][int]$MaximumCommandLineCharacters = 30000,
        [Threading.CancellationToken]$CancellationToken = [Threading.CancellationToken]::None,
        [AllowEmptyCollection()][string[]]$RemoveEnvironmentVariables = @()
    )

    $timer = [Diagnostics.Stopwatch]::StartNew()
    $process = $null
    $stdout = $null
    $stderr = $null
    $exitCode = $null
    $processId = $null
    $started = $false
    $timedOut = $false
    $cancelled = $false
    $launchError = $null
    $terminationError = $null
    $terminationAttempted = $false
    $captureErrors = New-Object 'System.Collections.Generic.List[string]'
    $renderedArguments = ''
    $resolvedExecutable = $Executable
    try {
        try {
            $serialized = ConvertTo-NativeArgumentString -Arguments $Arguments
            $renderedArguments = ConvertTo-NativeLogText -Text $serialized
            if ([string]::IsNullOrWhiteSpace($Executable) -or
                $Executable.IndexOf([char]0) -ge 0 -or
                -not [IO.Path]::IsPathRooted($Executable) -or
                [IO.Path]::GetFullPath($Executable) -ine $Executable -or
                [IO.Path]::GetExtension($Executable) -ine '.exe') {
                throw 'Native executable must be an explicitly resolved absolute .exe path.'
            }
            $resolved = @(Resolve-Path -LiteralPath $Executable -ErrorAction Stop)
            if ($resolved.Count -ne 1 -or $resolved[0].Provider.Name -ne 'FileSystem') {
                throw 'Native executable must resolve to exactly one FileSystem file.'
            }
            $file = Get-Item -LiteralPath $resolved[0].ProviderPath -Force -ErrorAction Stop
            if ($file -isnot [IO.FileInfo] -or $file.Extension -ine '.exe') {
                throw 'Native executable must resolve to an existing .exe file.'
            }
            $resolvedExecutable = $file.FullName
            $null = Assert-NativeCommandLength -Executable $resolvedExecutable -SerializedArguments $serialized `
                -MaximumCommandLineCharacters $MaximumCommandLineCharacters
            $startInfo = New-Object Diagnostics.ProcessStartInfo
            $startInfo.FileName = $resolvedExecutable
            $startInfo.Arguments = $serialized
            $startInfo.UseShellExecute = $false
            $startInfo.CreateNoWindow = $true
            $startInfo.RedirectStandardInput = $true
            $startInfo.RedirectStandardOutput = $true
            $startInfo.RedirectStandardError = $true
            $startInfo.StandardOutputEncoding = [Text.Encoding]::UTF8
            $startInfo.StandardErrorEncoding = [Text.Encoding]::UTF8
            foreach ($name in $RemoveEnvironmentVariables) {
                if ([string]::IsNullOrWhiteSpace($name) -or $name.IndexOfAny([char[]]@([char]0, [char]61)) -ge 0) {
                    throw 'Removed child environment variable names must be nonempty and contain neither NUL nor equals.'
                }
                $startInfo.EnvironmentVariables.Remove($name)
            }
            if ($CancellationToken.IsCancellationRequested) {
                $cancelled = $true
            } else {
                $process = New-Object Diagnostics.Process
                $process.StartInfo = $startInfo
                $started = $process.Start()
                if (-not $started) { throw 'Native process did not start.' }
                $processId = $process.Id
                $process.StandardInput.Close()
                # Start both asynchronous reads before any exit wait.
                $stdout = New-NativeStreamCapture -Reader $process.StandardOutput
                $stderr = New-NativeStreamCapture -Reader $process.StandardError
                while (-not $process.HasExited) {
                    if ($CancellationToken.IsCancellationRequested) { $cancelled = $true; break }
                    if ($timer.ElapsedMilliseconds -ge $TimeoutMilliseconds) { $timedOut = $true; break }
                    $stdoutReady = Receive-NativeStreamCapture -State $stdout -MaximumCaptureCharacters $MaximumCaptureCharacters
                    $stderrReady = Receive-NativeStreamCapture -State $stderr -MaximumCaptureCharacters $MaximumCaptureCharacters
                    if ($stdout.Error -or $stderr.Error) { break }
                    if (-not ($stdoutReady -or $stderrReady)) {
                        $remaining = $TimeoutMilliseconds - $timer.ElapsedMilliseconds
                        if ($remaining -gt 0) { $null = $process.WaitForExit([int][Math]::Min(20, $remaining)) }
                    }
                }
            }
        } catch {
            if ($started) { $captureErrors.Add($_.Exception.Message) } else { $launchError = $_.Exception.Message }
        }

        if ($started) {
            try {
                if (-not $process.HasExited) {
                    $terminationAttempted = $true
                    $terminationError = Stop-OwnedNativeProcess -Process $process -TimeoutMilliseconds $TerminationTimeoutMilliseconds
                }
            } catch {
                $terminationError = 'Owned native process termination failed; cleanup was best effort. ' + $_.Exception.Message
            }
            $captureTimer = [Diagnostics.Stopwatch]::StartNew()
            try {
                while (($null -ne $stdout -and -not $stdout.Closed) -or ($null -ne $stderr -and -not $stderr.Closed)) {
                    if ($captureTimer.ElapsedMilliseconds -ge $CaptureTimeoutMilliseconds) {
                        $captureErrors.Add("Native process stream capture timed out after $CaptureTimeoutMilliseconds ms; inherited stream handles may still be open. Captured output may be incomplete.")
                        break
                    }
                    $stdoutReady = Receive-NativeStreamCapture -State $stdout -MaximumCaptureCharacters $MaximumCaptureCharacters
                    $stderrReady = Receive-NativeStreamCapture -State $stderr -MaximumCaptureCharacters $MaximumCaptureCharacters
                    if (-not ($stdoutReady -or $stderrReady)) { [Threading.Thread]::Sleep(10) }
                }
                if ($process.HasExited) { $exitCode = $process.ExitCode }
            } catch {
                $captureErrors.Add($_.Exception.Message)
            } finally {
                $captureTimer.Stop()
            }
        }
    } finally {
        # Also protect the immediate child if an unexpected exception interrupts
        # lifecycle/capture work. A failed stop is attempted once, never retried
        # with an unbounded wait or an image-name-wide operation.
        if ($started -and -not $terminationAttempted) {
            try {
                if (-not $process.HasExited) {
                    $terminationAttempted = $true
                    $terminationError = Stop-OwnedNativeProcess -Process $process -TimeoutMilliseconds $TerminationTimeoutMilliseconds
                }
            } catch {
                $terminationError = 'Owned native process termination failed; cleanup was best effort. ' + $_.Exception.Message
            }
        }
        foreach ($stream in @($stdout, $stderr)) {
            if ($null -eq $stream) { continue }
            if ($stream.Error) { $captureErrors.Add($stream.Error) }
            if ($stream.Truncated) {
                $name = if ($stream -eq $stdout) { 'stdout' } else { 'stderr' }
                $captureErrors.Add("Native $name capture exceeded $MaximumCaptureCharacters characters; retained output was truncated while the stream was drained.")
            }
            try { $stream.Reader.Dispose() } catch { $captureErrors.Add($_.Exception.Message) }
            # Observe already completed faults without waiting for inherited pipes.
            if ($null -ne $stream.PendingRead -and $stream.PendingRead.IsFaulted) { $null = $stream.PendingRead.Exception }
        }
        if ($null -ne $process) {
            try { $process.Dispose() } catch { $captureErrors.Add($_.Exception.Message) }
        }
        $timer.Stop()
    }
    $captureError = if ($captureErrors.Count -gt 0) { $captureErrors -join ' ' } else { $null }
    return [pscustomobject]@{
        Executable = $resolvedExecutable
        RenderedArguments = $renderedArguments
        ExitCode = $exitCode
        Stdout = if ($null -ne $stdout) { $stdout.Text.ToString() } else { '' }
        Stderr = if ($null -ne $stderr) { $stderr.Text.ToString() } else { '' }
        ElapsedMilliseconds = $timer.ElapsedMilliseconds
        ProcessId = $processId
        Started = $started
        TimedOut = $timedOut
        Cancelled = $cancelled
        LaunchError = $launchError
        CaptureError = $captureError
        TerminationError = $terminationError
        StdoutTruncated = ($null -ne $stdout -and $stdout.Truncated)
        StderrTruncated = ($null -ne $stderr -and $stderr.Truncated)
        Succeeded = ($started -and $exitCode -eq 0 -and -not $timedOut -and -not $cancelled -and
            -not $launchError -and -not $captureError -and -not $terminationError)
    }
}

function Write-NativeProcessLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]$Result,
        [Parameter(Mandatory=$true)][string]$LiteralPath,
        [string]$Label = 'Native process'
    )

    $safeLabel = ConvertTo-NativeLogText -Text $Label
    @(
        "$safeLabel executable: $(ConvertTo-NativeLogText -Text $Result.Executable)"
        "$safeLabel arguments: $($Result.RenderedArguments)"
        "$safeLabel exit: $($Result.ExitCode); elapsed: $($Result.ElapsedMilliseconds) ms; PID: $($Result.ProcessId)"
        "$safeLabel started: $($Result.Started); timed out: $($Result.TimedOut); cancelled: $($Result.Cancelled); succeeded: $($Result.Succeeded)"
        "$safeLabel launch error: $(ConvertTo-NativeLogText -Text $Result.LaunchError)"
        "$safeLabel capture error: $(ConvertTo-NativeLogText -Text $Result.CaptureError)"
        "$safeLabel termination error: $(ConvertTo-NativeLogText -Text $Result.TerminationError)"
        "$safeLabel stdout truncated: $($Result.StdoutTruncated); stderr truncated: $($Result.StderrTruncated)"
        "$safeLabel stdout:"
        $Result.Stdout
        "$safeLabel stderr:"
        $Result.Stderr
    ) | Write-RunLog -LiteralPath $LiteralPath -Append
}

function Invoke-PdfToolJob {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][ValidateSet('Pdftk', 'Ghostscript')][string]$Tool,
        [Parameter(Mandatory=$true)][string]$Executable,
        [Parameter(Mandatory=$true)][AllowEmptyCollection()][object[]]$InputPaths,
        [Parameter(Mandatory=$true)][string]$OutputPath,
        [ValidateRange(1, 2147483647)][int]$TimeoutMilliseconds = 900000
    )

    $native = $null
    $outputError = $null
    $cleanupError = $null
    $published = $false
    $ownedDirectory = $null
    $stagedOutput = $null
    try {
        if ($InputPaths.Count -eq 0 -or ($Tool -eq 'Ghostscript' -and $InputPaths.Count -ne 1)) {
            throw 'PDFtk requires at least one input; Ghostscript requires exactly one master input.'
        }
        foreach ($path in @($InputPaths) + @($OutputPath)) {
            if ($path -isnot [string] -or [string]::IsNullOrWhiteSpace($path) -or
                -not [IO.Path]::IsPathRooted($path) -or [IO.Path]::GetFullPath($path) -ine $path) {
                throw 'PDF tool input/output paths must be literal absolute file paths.'
            }
            if ($path.Length -ge 260) {
                throw "Unsupported PDF tool path '$path': this workflow limits file paths to fewer than 260 UTF-16 characters. Use shorter folders; source files will not be renamed."
            }
        }
        foreach ($inputPath in $InputPaths) {
            if (-not [IO.File]::Exists($inputPath)) { throw "PDF input is missing or inaccessible: '$inputPath'." }
            if ($inputPath -ieq $OutputPath) { throw 'PDF output must be separate from every source file.' }
        }
        if (Test-Path -LiteralPath $OutputPath) { throw "PDF output already exists: '$OutputPath'. Choose a fresh output; existing files are never overwritten." }
        $parent = [IO.Path]::GetDirectoryName($OutputPath)
        if (-not [IO.Directory]::Exists($parent)) { throw "PDF output directory does not exist: '$parent'." }
        $candidate = Join-Path $parent ('.WinPDFMerge_' + [Guid]::NewGuid().ToString('N') + '.tmp')
        $stagedOutput = Join-Path $candidate 'output.pdf'
        if ($stagedOutput.Length -ge 260) {
            throw 'Output folder is too long for a private native output path (260-character limit). Use a shorter output folder.'
        }
        # New-Item without Force refuses an existing directory. Only after that
        # succeeds do we own this exact directory and its one known output file.
        $null = New-Item -ItemType Directory -Path $candidate -ErrorAction Stop
        $ownedDirectory = $candidate
        if ($Tool -eq 'Pdftk') {
            $arguments = @($InputPaths) + @('cat', 'output', $stagedOutput, 'compress', 'dont_ask')
            $removeEnvironment = @()
        } else {
            # GS otherwise can return zero and write a blank PDF after a PDF
            # interpreter error. Signal that error via its native exit status;
            # structural/page-total validation is still required separately.
            $arguments = @('-dBATCH', '-dNOPAUSE', '-dSAFER', '-dPDFSTOPONERROR', '-sDEVICE=pdfwrite',
                '-dCompatibilityLevel=1.6', '-dPDFSETTINGS=/screen', '-dDetectDuplicateImages=true',
                '-o', $stagedOutput, '-f', $InputPaths[0])
            $removeEnvironment = @('GS_OPTIONS')
        }
        $native = Invoke-NativeProcess -Executable $Executable -Arguments $arguments `
            -TimeoutMilliseconds $TimeoutMilliseconds -RemoveEnvironmentVariables $removeEnvironment
        if (-not $native.Succeeded) {
            throw "$Tool failed. Check native exit/launch/capture/timeout details and both streams in the log. The backend may reject Unicode or long paths; source files were not renamed."
        }
        if (-not [IO.File]::Exists($stagedOutput) -or (Get-Item -LiteralPath $stagedOutput).Length -eq 0) {
            throw "$Tool did not produce a nonempty private output."
        }
        # File.Move refuses an existing target, including a collision after the
        # preflight. Structural/page-total validation is a separate later gate.
        [IO.File]::Move($stagedOutput, $OutputPath)
        $published = $true
    } catch {
        $outputError = $_.Exception.Message
    } finally {
        if ($null -ne $ownedDirectory) {
            try {
                # Remove only the known run-owned file, then the empty directory.
                if ([IO.File]::Exists($stagedOutput)) { [IO.File]::Delete($stagedOutput) }
                [IO.Directory]::Delete($ownedDirectory, $false)
            } catch { $cleanupError = 'Private native output cleanup was best effort: ' + $_.Exception.Message }
        }
    }
    return [pscustomobject]@{
        NativeResult = $native
        OutputPath = $OutputPath
        OutputPublished = $published
        OutputError = $outputError
        CleanupError = $cleanupError
        Succeeded = ($published -and -not $outputError)
    }
}

function Get-DependencyExecutablePath {
    param([string]$Path, [string]$ExpectedName)

    if ([string]::IsNullOrWhiteSpace($Path) -or
        [IO.Path]::GetFileName($Path) -ine $ExpectedName) { return $null }
    try {
        $resolved = @(Resolve-Path -LiteralPath $Path -ErrorAction Stop)
        if ($resolved.Count -ne 1 -or $resolved[0].Provider.Name -ne 'FileSystem') { return $null }
        $file = Get-Item -LiteralPath $resolved[0].ProviderPath -Force -ErrorAction Stop
        if ($file -isnot [IO.FileInfo] -or $file.Name -ine $ExpectedName) { return $null }
        return $file.FullName
    } catch {
        return $null
    }
}

function Find-PathApplication {
    param([string]$Name)

    $applications = @(Get-Command -Name $Name -CommandType Application -All -ErrorAction SilentlyContinue)
    foreach ($application in $applications) {
        if ($application.CommandType -ne [Management.Automation.CommandTypes]::Application) { continue }
        $path = Get-DependencyExecutablePath -Path $application.Path -ExpectedName $Name
        if ($path) { return $path }
    }
    return $null
}

function Find-Pdftk {
    $path = Find-PathApplication -Name 'pdftk.exe'
    if ($path) { return $path }
    # Preserve the existing common-location priority; add the x86 Server path.
    $locations = @(
        @{ Root = $Env:ProgramFiles; Relative = 'PDFtk Server\bin\pdftk.exe' },
        @{ Root = ${Env:ProgramFiles(x86)}; Relative = 'PDFtk\bin\pdftk.exe' },
        @{ Root = ${Env:ProgramFiles(x86)}; Relative = 'PDFtk Server\bin\pdftk.exe' }
    )
    foreach ($location in $locations) {
        if ([string]::IsNullOrWhiteSpace($location.Root)) { continue }
        $path = Get-DependencyExecutablePath -Path (Join-Path $location.Root $location.Relative) -ExpectedName 'pdftk.exe'
        if ($path) { return $path }
    }
    return $null
}

function Find-Ghostscript {
    foreach ($name in @('gswin64c.exe', 'gswin32c.exe')) {
        $path = Find-PathApplication -Name $name
        if ($path) { return $path }
    }
    $installations = New-Object 'System.Collections.Generic.List[object]'
    $roots = @($Env:ProgramFiles, ${Env:ProgramFiles(x86)})
    for ($priority = 0; $priority -lt $roots.Count; $priority++) {
        if ([string]::IsNullOrWhiteSpace($roots[$priority])) { continue }
        $root = Join-Path $roots[$priority] 'gs'
        foreach ($directory in @(Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue)) {
            $match = [regex]::Match($directory.Name, '^gs([0-9]+(?:\.[0-9]+){1,3})$')
            [version]$version = $null
            if (-not $match.Success -or -not [version]::TryParse($match.Groups[1].Value, [ref]$version)) { continue }
            $installations.Add([pscustomobject]@{ Version = $version; Priority = $priority; Directory = $directory.FullName })
        }
    }
    $installations.Sort([System.Comparison[object]]{
        param($left, $right)
        $comparison = $right.Version.CompareTo($left.Version)
        if ($comparison -eq 0) { $comparison = $left.Priority.CompareTo($right.Priority) }
        if ($comparison -eq 0) { $comparison = [string]::CompareOrdinal($left.Directory, $right.Directory) }
        return $comparison
    })
    foreach ($installation in $installations) {
        foreach ($name in @('gswin64c.exe', 'gswin32c.exe')) {
            $path = Get-DependencyExecutablePath -Path (Join-Path $installation.Directory ('bin\' + $name)) -ExpectedName $name
            if ($path) { return $path }
        }
    }
    return $null
}

function Invoke-DependencyVersionProbe {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [ValidateRange(1, 60000)][int]$TimeoutMilliseconds = 5000
    )

    $result = Invoke-NativeProcess -Executable $Path -Arguments @('--version') `
        -TimeoutMilliseconds $TimeoutMilliseconds -RemoveEnvironmentVariables @('GS_OPTIONS')
    if ($result.TerminationError) { Write-Warning $result.TerminationError }
    if ($result.LaunchError) { throw $result.LaunchError }
    if ($result.TimedOut) { throw "Version probe timed out after $TimeoutMilliseconds ms." }
    if ($result.Cancelled) { throw 'Version probe was cancelled.' }
    if ($result.CaptureError) { throw ('Version probe stream capture failed. ' + $result.CaptureError) }
    if ($result.TerminationError) { throw $result.TerminationError }
    # Nonzero version exits remain the existing caller's useful tool-specific
    # diagnostic; version parsing consumes both streams exactly as before.
    return $result
}

function Get-NativeToolVersion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][ValidateSet('PdfTk', 'Ghostscript')][string]$Tool
    )

    $probe = Invoke-DependencyVersionProbe -Path $Path
    $diagnostic = 'stdout: {0}; stderr: {1}' -f $probe.Stdout.Trim(), $probe.Stderr.Trim()
    if ($diagnostic.Length -gt 2048) { $diagnostic = $diagnostic.Substring(0, 2048) + ' [truncated]' }
    if ($probe.ExitCode -ne 0) { throw "$Tool version probe failed (exit $($probe.ExitCode)). $diagnostic" }
    $pattern = if ($Tool -eq 'PdfTk') { '(?m)^pdftk ([0-9]+(?:\.[0-9]+){1,3})(?=\s|$)' } else { '(?m)^([0-9]+(?:\.[0-9]+){1,3})\s*$' }
    foreach ($output in @($probe.Stdout, $probe.Stderr)) {
        $match = [regex]::Match($output, $pattern)
        [version]$version = $null
        if ($match.Success -and [version]::TryParse($match.Groups[1].Value, [ref]$version)) {
            # Preserve actual spelling such as PDFtk 2.02, not normalized 2.2.
            return $match.Groups[1].Value
        }
    }
    throw "$Tool version probe returned unrecognized output. $diagnostic"
}

function Compare-NaturalName {
    param([string]$Left, [string]$Right)

    # Non-ASCII digits are text. Compare maximal runs without numeric parsing.
    $leftRuns = [regex]::Matches($Left, '[0-9]+|[^0-9]+')
    $rightRuns = [regex]::Matches($Right, '[0-9]+|[^0-9]+')
    $runCount = [Math]::Min($leftRuns.Count, $rightRuns.Count)
    for ($index = 0; $index -lt $runCount; $index++) {
        $leftRun = $leftRuns[$index].Value
        $rightRun = $rightRuns[$index].Value
        $leftIsNumber = $leftRun[0] -ge [char]'0' -and $leftRun[0] -le [char]'9'
        $rightIsNumber = $rightRun[0] -ge [char]'0' -and $rightRun[0] -le [char]'9'
        if ($leftIsNumber -and $rightIsNumber) {
            $leftDigits = $leftRun.TrimStart([char[]]'0')
            $rightDigits = $rightRun.TrimStart([char[]]'0')
            $comparison = $leftDigits.Length.CompareTo($rightDigits.Length)
            if ($comparison -eq 0) {
                $comparison = [string]::CompareOrdinal($leftDigits, $rightDigits)
            }
            if ($comparison -eq 0) {
                # Resolve an equal numeric run before considering later runs.
                $comparison = $leftRun.Length.CompareTo($rightRun.Length)
            }
        } else {
            # Also defines the mixed digit/text rule: ordinal text comparison.
            $comparison = [string]::Compare($leftRun, $rightRun, [StringComparison]::OrdinalIgnoreCase)
        }
        if ($comparison -ne 0) { return $comparison }
    }
    return $leftRuns.Count.CompareTo($rightRuns.Count)
}

function Compare-PdfInput {
    param($Left, $Right)

    $comparison = Compare-NaturalName -Left $Left.BaseName -Right $Right.BaseName
    if ($comparison -eq 0) {
        # Case differences are deferred until all natural segments compare equal.
        $comparison = [string]::CompareOrdinal($Left.BaseName, $Right.BaseName)
    }
    if ($comparison -eq 0) {
        # Discovery supplies canonical absolute FileInfo.FullName values.
        $comparison = [string]::CompareOrdinal($Left.FullName, $Right.FullName)
    }
    return $comparison
}

function Sort-PdfInputs {
    param([object[]]$Inputs)

    # Sort a separate collection, retaining the frozen FileInfo objects.
    $ordered = New-Object 'System.Collections.Generic.List[object]'
    if ($Inputs.Count -gt 0) { $ordered.AddRange($Inputs) }
    $ordered.Sort([System.Comparison[object]]{
        param($left, $right)
        Compare-PdfInput -Left $left -Right $right
    })
    return $ordered
}
function Sanitize-FileName([string]$name) {
    $invalid = [IO.Path]::GetInvalidFileNameChars() -join ''
    $re = "[{0}]" -f ([Regex]::Escape($invalid))
    ($name -replace $re, '_').Trim()
}
