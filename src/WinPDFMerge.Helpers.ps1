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

function Resolve-OutputDirectory {
    [CmdletBinding()]
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw 'OutputFolder must name one existing writable FileSystem directory. Create the directory first and supply -OutputFolder.'
    }
    if ($Path.IndexOfAny([char[]]'*?') -ge 0) {
        throw "OutputFolder does not support wildcard expansion: '$Path'. Supply one literal directory."
    }
    try { $resolved = @(Resolve-Path -LiteralPath $Path -ErrorAction Stop) }
    catch { throw "OutputFolder must be an existing accessible directory: '$Path'. Create it first or choose another -OutputFolder. $($_.Exception.Message)" }
    if ($resolved.Count -ne 1 -or $resolved[0].Provider.Name -ne 'FileSystem') {
        throw "OutputFolder must resolve to exactly one FileSystem directory: '$Path'."
    }
    $directory = Get-Item -LiteralPath $resolved[0].ProviderPath -Force -ErrorAction Stop
    if (-not $directory.PSIsContainer) { throw "OutputFolder is not a directory: '$Path'." }
    $fullPath = [IO.Path]::GetFullPath($directory.FullName)
    $root = [IO.Path]::GetPathRoot($fullPath)
    if ($fullPath.Length -gt $root.Length) { $fullPath = $fullPath.TrimEnd([char[]]'\/') }
    return $fullPath
}

function Assert-MergeDirectoryPath {
    param([string]$Path, [string]$Role)

    # Reject a junction/symlink/mount point at any component, including an
    # ancestor of an ordinary leaf. Do not guess a reparse target's identity.
    $current = New-Object IO.DirectoryInfo($Path)
    while ($null -ne $current) {
        $item = Get-Item -LiteralPath $current.FullName -Force -ErrorAction Stop
        if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "$Role contains an unsupported junction or reparse directory: '$($current.FullName)'. Choose a direct directory path for SourceFolder and -OutputFolder."
        }
        $current = $current.Parent
    }
}

function Get-MergeDirectoryIdentity {
    param([string]$Path)

    if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) {
        throw 'Directory identity requires Windows; choose a supported Windows filesystem.'
    }
    # Lazy, small Windows metadata adapter. Importing helpers compiles nothing
    # and performs no filesystem/native work.128-bit IDs avoid truncating ReFS
    # identity, but only the actually tested filesystem is a support claim.
    if (-not ('WinPDFMerger.DirectoryIdentity' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
namespace WinPDFMerger {
    public static class DirectoryIdentity {
        [StructLayout(LayoutKind.Sequential)]
        private struct FileIdInfo {
            public ulong VolumeSerialNumber;
            public ulong FileIdLow;
            public ulong FileIdHigh;
        }
        [DllImport("kernel32.dll", CharSet=CharSet.Unicode, ExactSpelling=true, SetLastError=true)]
        private static extern SafeFileHandle CreateFileW(string name, uint access, uint share,
            IntPtr security, uint creation, uint flags, IntPtr template);
        [DllImport("kernel32.dll", ExactSpelling=true, SetLastError=true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetFileInformationByHandleEx(SafeFileHandle handle, int informationClass,
            out FileIdInfo information, uint bufferSize);
        public static string Get(string path) {
            // Metadata only, all sharing modes, OPEN_EXISTING/BACKUP_SEMANTICS.
            using (SafeFileHandle handle = CreateFileW(path, 0, 7, IntPtr.Zero, 3, 0x02000000, IntPtr.Zero)) {
                if (handle.IsInvalid) throw new IOException("Cannot open directory identity: " + path,
                    new Win32Exception(Marshal.GetLastWin32Error()));
                FileIdInfo info;
                // FileIdInfo=18; FILE_ID_INFO is volume64 followed by opaque ID128.
                if (!GetFileInformationByHandleEx(handle, 18, out info, (uint)Marshal.SizeOf(typeof(FileIdInfo))))
                    throw new IOException("Cannot establish directory identity: " + path,
                        new Win32Exception(Marshal.GetLastWin32Error()));
                if (info.FileIdLow == 0 && info.FileIdHigh == 0)
                    throw new IOException("Ambiguous directory identity: " + path);
                return info.VolumeSerialNumber.ToString("X16", CultureInfo.InvariantCulture) + ":" +
                    info.FileIdHigh.ToString("X16", CultureInfo.InvariantCulture) +
                    info.FileIdLow.ToString("X16", CultureInfo.InvariantCulture);
            }
        }
    }
}
'@ -ErrorAction Stop
    }
    return [WinPDFMerger.DirectoryIdentity]::Get($Path)
}

function Assert-MergeDirectories {
    param([string]$SourceFolder, [string]$OutputFolder)

    Assert-MergeDirectoryPath -Path $SourceFolder -Role SourceFolder
    Assert-MergeDirectoryPath -Path $OutputFolder -Role OutputFolder
    $sourceIdentity = Get-MergeDirectoryIdentity -Path $SourceFolder
    $outputIdentity = Get-MergeDirectoryIdentity -Path $OutputFolder
    if ($sourceIdentity -ceq $outputIdentity) {
        throw 'SourceFolder and OutputFolder refer to the same directory. Choose a separate existing -OutputFolder; sources cannot also be the destination.'
    }
}

function Test-OutputDirectoryWritable {
    param([string]$OutputFolder)

    $probe = [IO.Path]::Combine($OutputFolder, ('.WinPDFMerge_probe_' + [Guid]::NewGuid().ToString('N') + '.tmp'))
    $stream = $null
    $failure = $null
    try {
        # CreateNew cannot overwrite a foreign candidate. DeleteOnClose belongs
        # only to the successfully opened handle; no name-based cleanup sweep.
        $stream = [IO.FileStream]::new($probe, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write,
            [IO.FileShare]::None, 4096, [IO.FileOptions]::DeleteOnClose)
        $stream.WriteByte(0)
        $stream.Flush()
    } catch { $failure = $_.Exception.Message }
    finally {
        if ($null -ne $stream) {
            try { $stream.Dispose() }
            catch { $failure = "Owned writability probe cleanup failed: $($_.Exception.Message)" }
        }
    }
    if ($failure) {
        throw "OutputFolder '$OutputFolder' is not writable or its owned probe could not be cleaned. Choose an existing writable -OutputFolder. $failure"
    }
}

function New-MergeRunIdentity {
    [CmdletBinding()]
    param(
        [AllowEmptyString()][string]$SourceFolder,
        [Parameter(Mandatory=$true)][string]$OutputFolder,
        [datetime]$Timestamp = [datetime]::Now,
        [ValidatePattern('^[0-9a-fA-F]{16}$')][string]$RunSuffix = ([Guid]::NewGuid().ToString('N').Substring(0, 16))
    )

    $stamp = $Timestamp.ToString('yyyyMMdd_HHmmss', [Globalization.CultureInfo]::InvariantCulture)
    $suffix = $RunSuffix.ToLowerInvariant()
    $leaf = ''
    if (-not [string]::IsNullOrWhiteSpace($SourceFolder)) {
        $source = $SourceFolder.TrimEnd([char[]]'\/')
        $root = [IO.Path]::GetPathRoot($SourceFolder).TrimEnd([char[]]'\/')
        if ($source -ine $root) { $leaf = [IO.Path]::GetFileName($source) }
    }
    $label = (Sanitize-FileName $leaf).TrimEnd([char[]]' .')
    if ([string]::IsNullOrWhiteSpace($label)) { $label = 'root' }
    # Include the private output layout in preflight, before probe,
    # log or native work. Existing native backend operands stay below260.
    $stage = [IO.Path]::Combine([IO.Path]::Combine($OutputFolder, ('.WinPDFMerge_' + ('0' * 32) + '.tmp')), 'output.pdf')
    $fixedEmail = [IO.Path]::Combine($OutputFolder, ('WinPDFMerge__' + $stamp + '_' + $suffix + '_email.pdf'))
    $labelLimit = [Math]::Min(64, (259 - $fixedEmail.Length))
    if ($stage.Length -ge 260 -or $labelLimit -lt 1) {
        throw 'Output paths have insufficient room for safe names and private native output. Choose a shorter existing -OutputFolder.'
    }
    if ($label.Length -gt $labelLimit) {
        $label = $label.Substring(0, $labelLimit)
        if ([char]::IsHighSurrogate($label[$label.Length - 1])) { $label = $label.Substring(0, $label.Length - 1) }
        $label = $label.TrimEnd([char[]]' .')
        if ([string]::IsNullOrEmpty($label)) { $label = 'r' }
    }
    $baseName = 'WinPDFMerge_' + $label + '_' + $stamp + '_' + $suffix
    return [pscustomobject]@{
        OutputFolder = $OutputFolder; FolderLabel = $label; Timestamp = $stamp; RunSuffix = $suffix
        BaseName = $baseName
        MasterPath = [IO.Path]::Combine($OutputFolder, ($baseName + '.pdf'))
        EmailPath = [IO.Path]::Combine($OutputFolder, ($baseName + '_email.pdf'))
        LogPath = [IO.Path]::Combine($OutputFolder, ($baseName + '.log'))
    }
}

function Reserve-MergeRunIdentity {
    param($Identity)

    foreach ($path in @($Identity.MasterPath, $Identity.EmailPath, $Identity.LogPath)) {
        if ([IO.File]::Exists($path) -or [IO.Directory]::Exists($path)) {
            throw "Run identity already exists at '$path'. No existing output was replaced; run again for a fresh identity."
        }
    }
    $stream = $null
    try {
        # Atomically claim the cooperating run's identity via its log. All entry
        # writes append afterward; final PDF moves still enforce no-overwrite.
        $stream = [IO.FileStream]::new($Identity.LogPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::Read)
    } catch { throw "Cannot reserve run identity in OutputFolder '$($Identity.OutputFolder)'. No existing file was replaced. $($_.Exception.Message)" }
    finally { if ($null -ne $stream) { $stream.Dispose() } }
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

function ConvertFrom-PdfDocumentData {
    [CmdletBinding()]
    param([AllowEmptyString()][string]$Text)

    # Only a physical stdout label supplies the count. Horizontal whitespace
    # cannot consume another line, and a malformed duplicate is still ambiguous.
    $labels = @([regex]::Matches($Text, '(?m)^NumberOfPages:[^\r\n]*\r?$'))
    if ($labels.Count -ne 1) { throw 'PDF document data must contain exactly one labeled page count.' }
    $value = [regex]::Match($labels[0].Value, '^NumberOfPages:[ \t]*([0-9]+)[ \t]*\r?$')
    $count = [long]0
    if (-not $value.Success -or -not [long]::TryParse($value.Groups[1].Value,
            [Globalization.NumberStyles]::None, [Globalization.CultureInfo]::InvariantCulture, [ref]$count) -or $count -le 0) {
        throw 'PDF document data has an invalid, zero or unsupported page count.'
    }
    return $count
}

function Get-PdfInputSnapshot {
    [CmdletBinding()]
    param([Parameter(Mandatory=$true)][string]$LiteralPath)

    try {
        if ([string]::IsNullOrWhiteSpace($LiteralPath) -or $LiteralPath.IndexOf([char]0) -ge 0 -or
            -not [IO.Path]::IsPathRooted($LiteralPath) -or [IO.Path]::GetFullPath($LiteralPath) -ine $LiteralPath) {
            throw 'Use a literal canonical absolute input file path.'
        }
        if ($LiteralPath.Length -ge 260) { throw 'Input file paths must be fewer than 260 UTF-16 characters; use shorter folders without renaming sources.' }
        # Get-Item reads fresh metadata; discovery FileInfo caches are not reused.
        $file = Get-Item -LiteralPath $LiteralPath -Force -ErrorAction Stop
        if ($file -isnot [IO.FileInfo]) { throw 'The input must be a filesystem file.' }
        if ($file.Length -eq 0) { throw 'The input is empty.' }
        return [pscustomobject]@{
            FullName = $file.FullName
            Length = [long]$file.Length
            LastWriteTimeUtcTicks = [long]$file.LastWriteTimeUtc.Ticks
        }
    } catch { throw "PDF input '$LiteralPath' cannot be inspected: $($_.Exception.Message)" }
}

function Assert-PdfInputSnapshot {
    param([Parameter(Mandatory=$true)]$Snapshot)

    $current = Get-PdfInputSnapshot -LiteralPath $Snapshot.FullName
    if ($current.Length -ne $Snapshot.Length -or $current.LastWriteTimeUtcTicks -ne $Snapshot.LastWriteTimeUtcTicks) {
        throw "PDF input '$($Snapshot.FullName)' changed during preflight. Use stable source documents and run again; no merge was started."
    }
}

function Read-PdfEnvelopeBytes {
    param([IO.Stream]$Stream, [long]$Offset, [ValidateRange(1, 8192)][int]$Count)

    $Stream.Position = $Offset
    $buffer = New-Object byte[] $Count
    $read = 0
    while ($read -lt $Count) {
        $received = $Stream.Read($buffer, $read, ($Count - $read))
        if ($received -eq 0) { throw 'PDF envelope changed or could not be read completely.' }
        $read += $received
    }
    return ,$buffer
}

function Assert-PdfInputEnvelope {
    [CmdletBinding()]
    param([Parameter(Mandatory=$true)][string]$LiteralPath)

    $stream = $null
    try {
        $stream = [IO.File]::Open($LiteralPath, [IO.FileMode]::Open, [IO.FileAccess]::Read,
            ([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
        $length = $stream.Length
        if ($length -eq 0) { throw 'The input is empty.' }
        # Latin1 preserves one byte per character; offsets are byte positions,
        # never UTF8 character positions or normalized newline positions.
        $encoding = [Text.Encoding]::GetEncoding(28591)
        $head = $encoding.GetString((Read-PdfEnvelopeBytes -Stream $stream -Offset 0 -Count ([int][Math]::Min(16, $length))))
        if ($head -cnotmatch '\A%PDF-[12]\.[0-9](?:\r\n|\r|\n)') { throw 'A supported PDF header must begin the input file.' }
        $tailLength = [int][Math]::Min(8192, $length)
        $tail = $encoding.GetString((Read-PdfEnvelopeBytes -Stream $stream -Offset ($length - $tailLength) -Count $tailLength))
        # Only the final footer counts. Earlier EOF/startxref0 records in valid
        # linearized/incremental PDFs do not decide this envelope check.
        $footer = [regex]::Match($tail, '(?:\A|[\r\n])[ \t]*startxref[ \t]*(?:\r\n|\r|\n)[ \t]*([0-9]+)[ \t]*(?:\r\n|\r|\n)%%EOF[\x00\t\n\f\r ]*\z')
        $offset = [long]0
        if (-not $footer.Success -or -not [long]::TryParse($footer.Groups[1].Value,
                [Globalization.NumberStyles]::None, [Globalization.CultureInfo]::InvariantCulture, [ref]$offset) -or
            $offset -le 0 -or $offset -ge $length) {
            throw 'PDF footer is unsupported or malformed: require final startxref with a positive in-file offset and terminal %%EOF within the last 8192 bytes.'
        }
        $target = $encoding.GetString((Read-PdfEnvelopeBytes -Stream $stream -Offset $offset -Count ([int][Math]::Min(1024, ($length - $offset)))))
        $table = $target -cmatch '\Axref(?:[\x00\t\n\f\r ]|\z)'
        $separator = '(?:[\x00\t\n\f\r ]|%[^\r\n]*(?:\r\n|\r|\n))+'
        $indirect = $target -cmatch ('\A[0-9]+' + $separator + '[0-9]+' + $separator + 'obj(?=[\x00\t\n\f\r ()<>\[\]{}/%])')
        if (-not $table -and -not $indirect) {
            throw 'PDF final cross-reference target is unsupported or malformed; expected xref or an indirect-object header within 1024 bytes.'
        }
        # This is an envelope plausibility guard. PDFtk still inspects the page
        # structure; neither check certifies every dictionary/stream or fidelity.
    } catch { throw "PDF input '$LiteralPath' failed envelope preflight: $($_.Exception.Message)" }
    finally { if ($null -ne $stream) { $stream.Dispose() } }
}

function Get-PdfDocumentInspection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$Executable,
        [Parameter(Mandatory=$true)][string]$LiteralPath,
        [ValidateRange(1, 2147483647)][int]$TimeoutMilliseconds = 900000
    )

    $native = $null
    $count = $null
    $inputError = $null
    try {
        $null = Get-PdfInputSnapshot -LiteralPath $LiteralPath
        Assert-PdfInputEnvelope -LiteralPath $LiteralPath
        # Read-only operation and explicit stdout target: no output PDF, password
        # workflow or repair operation. Shared runner closes stdin and bounds
        # execution, both streams, capture and immediate-owned termination.
        $native = Invoke-NativeProcess -Executable $Executable `
            -Arguments @($LiteralPath, 'dump_data_utf8', 'output', '-', 'dont_ask') -TimeoutMilliseconds $TimeoutMilliseconds
        if (-not $native.Succeeded) {
            throw "PDFtk document inspection failed (exit code: $($native.ExitCode)). See native launch/capture/timeout details and both streams; protected or unparseable inputs are unsupported without passwords or repair."
        }
        $count = ConvertFrom-PdfDocumentData -Text $native.Stdout
    } catch { $inputError = "PDF input '$LiteralPath' failed preflight: $($_.Exception.Message)" }
    return [pscustomobject]@{
        NativeResult = $native
        PageCount = $count
        InputError = $inputError
        Succeeded = ($null -eq $inputError)
    }
}

function Get-PdfInputInventory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$Executable,
        [Parameter(Mandatory=$true)][AllowEmptyCollection()][object[]]$Inputs,
        [string]$LogPath,
        [ValidateRange(1, 2147483647)][int]$TimeoutMilliseconds = 900000
    )

    if ($Inputs.Count -eq 0) { throw 'The page inventory requires at least one input file.' }
    $snapshots = New-Object 'System.Collections.Generic.List[object]'
    # Freeze every metadata record before inspecting the first input. Discovery
    # remains the sole enumeration; later files are never added to this run.
    foreach ($inputFile in $Inputs) {
        if ($inputFile -isnot [IO.FileInfo]) { throw 'The page inventory requires discovered filesystem input files.' }
        $snapshots.Add((Get-PdfInputSnapshot -LiteralPath $inputFile.FullName))
    }
    $entries = New-Object 'System.Collections.Generic.List[object]'
    $total = [long]0
    for ($index = 0; $index -lt $snapshots.Count; $index++) {
        $snapshot = $snapshots[$index]
        Assert-PdfInputSnapshot -Snapshot $snapshot
        $inspection = Get-PdfDocumentInspection -Executable $Executable -LiteralPath $snapshot.FullName -TimeoutMilliseconds $TimeoutMilliseconds
        if ($LogPath -and $null -ne $inspection.NativeResult) {
            # The logger echoes strings on the success stream. Keep raw evidence
            # in the file without mixing it into this function's result object.
            Write-NativeProcessLog -Result $inspection.NativeResult -LiteralPath $LogPath -Label ('Input preflight ' + ($index + 1)) | Out-Null
        }
        if (-not $inspection.Succeeded) { throw $inspection.InputError }
        Assert-PdfInputSnapshot -Snapshot $snapshot
        $pages = [long]$inspection.PageCount
        if ($pages -le 0 -or $total -gt ([long]::MaxValue - $pages)) {
            throw "Expected page total is zero or exceeds the supported limit at PDF input '$($snapshot.FullName)'."
        }
        $total += $pages
        $entries.Add([pscustomobject]@{
            FullName = $snapshot.FullName
            Length = $snapshot.Length
            LastWriteTimeUtcTicks = $snapshot.LastWriteTimeUtcTicks
            PageCount = $pages
        })
        # Each potentially large native result is logged promptly, then released;
        # only lightweight ordered metadata survives the inventory loop.
        $inspection = $null
    }
    return [pscustomobject]@{ Inputs = $entries.ToArray(); ExpectedPageCount = $total }
}

function Assert-PdfInputInventory {
    param([Parameter(Mandatory=$true)]$Inventory)
    foreach ($entry in $Inventory.Inputs) { Assert-PdfInputSnapshot -Snapshot $entry }
}

function New-PdfStaging {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$OutputFolder,
        [string]$RunIdentity = [Guid]::NewGuid().ToString('N'),
        [ValidatePattern('^[0-9a-fA-F]{32}$')][string]$StageSuffix = [Guid]::NewGuid().ToString('N')
    )

    $parent = Resolve-OutputDirectory -Path $OutputFolder
    Assert-MergeDirectoryPath -Path $parent -Role OutputFolder
    $parentIdentity = Get-MergeDirectoryIdentity -Path $parent
    $directory = [IO.Path]::Combine($parent, ('.WinPDFMerge_' + $StageSuffix.ToLowerInvariant() + '.tmp'))
    $master = [IO.Path]::Combine($directory, 'master.pdf')
    if ($master.Length -ge 260) { throw 'Output folder is too long for private native output (260-character limit). Choose a shorter -OutputFolder.' }
    # Directory.CreateDirectory accepts an existing directory. Use the Windows
    # create-new operation so a losing race never acquires cleanup ownership.
    if (-not ('WinPDFMerger.StagingDirectory' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
namespace WinPDFMerger {
    public static class StagingDirectory {
        [DllImport("kernel32.dll", CharSet=CharSet.Unicode, ExactSpelling=true, SetLastError=true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool CreateDirectoryW(string path, IntPtr security);
        public static void CreateNew(string path) {
            if (!CreateDirectoryW(path, IntPtr.Zero))
                throw new IOException("Cannot reserve new private staging directory: " + path,
                    new Win32Exception(Marshal.GetLastWin32Error()));
        }
    }
}
'@ -ErrorAction Stop
    }
    [WinPDFMerger.StagingDirectory]::CreateNew($directory)
    $marker = [IO.Path]::Combine($directory, 'owner.json')
    $stream = $null
    try {
        $directoryIdentity = Get-MergeDirectoryIdentity -Path $directory
        $text = [ordered]@{ SchemaVersion=1; RunIdentity=$RunIdentity; StageSuffix=$StageSuffix.ToLowerInvariant();
            CreatedUtc=[datetime]::UtcNow.ToString('o'); ProcessId=$PID; KnownFiles=@('master.pdf','email.pdf') } | ConvertTo-Json -Compress
        # Create and flush the marker, then retain a read handle without
        # delete/write sharing. Other readers can inspect ownership evidence;
        # a crash closes the handle but leaves the marker for manual inspection.
        $stream = [IO.File]::Open($marker, [IO.FileMode]::CreateNew, [IO.FileAccess]::ReadWrite, [IO.FileShare]::Read)
        $bytes = (New-Object Text.UTF8Encoding($false)).GetBytes($text)
        $stream.Write($bytes, 0, $bytes.Length)
        $stream.Flush()
        $stream.Dispose()
        $stream = [IO.File]::Open($marker, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
        if ([IO.File]::ReadAllText($marker, [Text.Encoding]::UTF8) -cne $text) { throw 'Ownership marker changed during initialization.' }
        return [pscustomobject]@{
            OutputFolder=$parent; OutputDirectoryIdentity=$parentIdentity
            DirectoryPath=$directory; DirectoryIdentity=$directoryIdentity
            MasterPath=$master; EmailPath=[IO.Path]::Combine($directory, 'email.pdf')
            MarkerPath=$marker; MarkerStream=$stream; MarkerText=$text; Cleaned=$false
        }
    } catch {
        $reason = $_.Exception.Message
        if ($null -ne $stream) { $stream.Dispose() }
        # A partly initialized directory is deliberately retained: cleanup has
        # not established its full ownership record. Never guess at its files.
        throw "Private staging initialization failed: $reason Staging may remain at '$directory'. Inspect it manually after all runs have stopped; no automatic orphan sweep is performed."
    }
}

function Assert-PdfStaging {
    param([Parameter(Mandatory=$true)]$Staging)

    if ($Staging.Cleaned -or $null -eq $Staging.MarkerStream -or -not $Staging.MarkerStream.CanRead) {
        throw 'Private staging ownership handle is no longer active.'
    }
    $parent = [IO.Path]::GetFullPath($Staging.OutputFolder)
    $directory = [IO.Path]::GetFullPath($Staging.DirectoryPath)
    if ($parent -cne $Staging.OutputFolder -or $directory -cne $Staging.DirectoryPath -or
        [IO.Path]::GetDirectoryName($directory) -ine $parent -or
        [IO.Path]::GetFileName($directory) -cnotmatch '^\.WinPDFMerge_[0-9a-f]{32}\.tmp$' -or
        $Staging.MasterPath -cne [IO.Path]::Combine($directory,'master.pdf') -or
        $Staging.EmailPath -cne [IO.Path]::Combine($directory,'email.pdf') -or
        $Staging.MarkerPath -cne [IO.Path]::Combine($directory,'owner.json')) {
        throw 'Private staging layout does not match the owned known paths.'
    }
    Assert-MergeDirectoryPath -Path $directory -Role 'Private staging'
    if ((Get-MergeDirectoryIdentity -Path $parent) -cne $Staging.OutputDirectoryIdentity -or
        (Get-MergeDirectoryIdentity -Path $directory) -cne $Staging.DirectoryIdentity) {
        throw 'Private staging directory identity changed; no files will be moved or cleaned.'
    }
    if ([IO.File]::ReadAllText($Staging.MarkerPath, [Text.Encoding]::UTF8) -cne $Staging.MarkerText) {
        throw 'Private staging ownership marker changed; no files will be moved or cleaned.'
    }
}

function Publish-PdfStagedOutput {
    param([Parameter(Mandatory=$true)]$Staging,
        [Parameter(Mandatory=$true)][string]$StagedPath,
        [Parameter(Mandatory=$true)][string]$OutputPath)

    Assert-PdfStaging -Staging $Staging
    if ($StagedPath -cne $Staging.MasterPath -and $StagedPath -cne $Staging.EmailPath) {
        throw 'Publication requires a known owned staged PDF path.'
    }
    if (-not [IO.Path]::IsPathRooted($OutputPath) -or [IO.Path]::GetFullPath($OutputPath) -cne $OutputPath -or
        [IO.Path]::GetDirectoryName($OutputPath) -ine $Staging.OutputFolder -or $OutputPath.Length -ge 260) {
        throw 'Publication requires a literal final path directly in the same output directory and volume.'
    }
    $file = Get-Item -LiteralPath $StagedPath -Force -ErrorAction Stop
    if ($file.PSIsContainer -or ($file.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0 -or $file.Length -eq 0) {
        throw 'Publication requires a nonempty regular owned staged PDF.'
    }
    # Two-argument File.Move never replaces an existing file. This operation,
    # not Test-Path, decides the collision outcome. Parent identity/reparse
    # checks keep the sibling move on the original destination volume.
    [IO.File]::Move($StagedPath, $OutputPath)
}

function Remove-PdfStaging {
    param([Parameter(Mandatory=$true)]$Staging)

    if ($Staging.Cleaned) { return [pscustomobject]@{ Cleaned=$true; CleanupError=$null; OrphanPath=$null } }
    $errorText = $null
    $markerDeleted = $false
    try {
        Assert-PdfStaging -Staging $Staging
        # Inspect only this owned directory, never scan for other run prefixes.
        # Unknown children/reparse files retain the marker and all contents.
        $known = @($Staging.MasterPath, $Staging.EmailPath, $Staging.MarkerPath)
        foreach ($child in @(Get-ChildItem -LiteralPath $Staging.DirectoryPath -Force -ErrorAction Stop)) {
            if ($child.FullName -cnotin $known -or $child.PSIsContainer -or
                ($child.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw 'Unexpected content in private staging; retained for manual inspection.'
            }
        }
        foreach ($path in @($Staging.MasterPath, $Staging.EmailPath)) {
            if ([IO.File]::Exists($path)) { [IO.File]::Delete($path) }
        }
        $Staging.MarkerStream.Dispose()
        [IO.File]::Delete($Staging.MarkerPath)
        $markerDeleted = $true
        [IO.Directory]::Delete($Staging.DirectoryPath, $false)
        $Staging.Cleaned = $true
    } catch {
        $errorText = "Private staging cleanup was best effort: $($_.Exception.Message) Staging remains at '$($Staging.DirectoryPath)'. Inspect it manually after all runs have stopped; no automatic orphan sweep is performed."
        # Directory removal can fail after deleting the marker (for example a
        # new unknown child). Restore evidence only in the same owned directory,
        # using CreateNew so an existing marker is never replaced.
        if ($markerDeleted -and -not [IO.File]::Exists($Staging.MarkerPath)) {
            $restore = $null
            try {
                Assert-MergeDirectoryPath -Path $Staging.DirectoryPath -Role 'Private staging'
                if ((Get-MergeDirectoryIdentity -Path $Staging.DirectoryPath) -cne $Staging.DirectoryIdentity -or
                    (Get-MergeDirectoryIdentity -Path $Staging.OutputFolder) -cne $Staging.OutputDirectoryIdentity) {
                    throw 'Original staging directory identity is unavailable.'
                }
                $restore = [IO.File]::Open($Staging.MarkerPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::Read)
                $bytes = (New-Object Text.UTF8Encoding($false)).GetBytes($Staging.MarkerText)
                $restore.Write($bytes, 0, $bytes.Length)
                $restore.Flush()
            } catch { $errorText += ' Ownership marker could not be retained: ' + $_.Exception.Message }
            finally { if ($null -ne $restore) { $restore.Dispose() } }
        }
    } finally {
        if ($null -ne $Staging.MarkerStream) { $Staging.MarkerStream.Dispose() }
    }
    return [pscustomobject]@{ Cleaned=$Staging.Cleaned; CleanupError=$errorText;
        OrphanPath=$(if ($Staging.Cleaned) { $null } else { $Staging.DirectoryPath }) }
}

function Invoke-PdfToolJob {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][ValidateSet('Pdftk', 'Ghostscript')][string]$Tool,
        [Parameter(Mandatory=$true)][string]$Executable,
        [Parameter(Mandatory=$true)][AllowEmptyCollection()][object[]]$InputPaths,
        [Parameter(Mandatory=$true)][string]$OutputPath,
        $Staging,
        [long]$ExpectedPageCount = 0,
        [string]$InspectionExecutable,
        [ValidateRange(1, 2147483647)][int]$TimeoutMilliseconds = 900000
    )

    $native = $null
    $outputError = $null
    $cleanupError = $null
    $validation = $null
    $validated = $false
    $validatedPages = $null
    $outputState = 'failed'
    $masterSnapshot = $null
    $masterBytes = $null
    $outputBytes = $null
    $published = $false
    $ownedStaging = $null
    $stagedOutput = $null
    try {
        if ($ExpectedPageCount -le 0) {
            throw "$Tool requires a positive frozen ExpectedPageCount; no native job was started."
        }
        if ($Tool -eq 'Ghostscript' -and [string]::IsNullOrWhiteSpace($InspectionExecutable)) {
            throw 'Ghostscript email jobs require the selected PDFtk InspectionExecutable; no native job was started.'
        }
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
        if ($Tool -eq 'Ghostscript') {
            if (-not [IO.Path]::IsPathRooted($InspectionExecutable) -or
                [IO.Path]::GetFullPath($InspectionExecutable) -ine $InspectionExecutable -or $InspectionExecutable.Length -ge 260) {
                throw 'Selected PDFtk InspectionExecutable must be a literal absolute path shorter than 260 UTF-16 characters.'
            }
            if (-not [IO.File]::Exists($InspectionExecutable)) { throw 'Selected PDFtk InspectionExecutable is missing; no email job was started.' }
            $masterFile = Get-Item -LiteralPath $InputPaths[0] -Force -ErrorAction Stop
            if ($masterFile -isnot [IO.FileInfo] -or ($masterFile.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw 'Email conversion requires a regular published master.'
            }
            $masterSnapshot = Get-PdfInputSnapshot -LiteralPath $InputPaths[0]
            if ($masterSnapshot.Length -le 0) { throw 'Email conversion requires a nonempty published master.' }
            $masterBytes = $masterSnapshot.Length
        }
        if (Test-Path -LiteralPath $OutputPath) { throw "PDF output already exists: '$OutputPath'. Choose a fresh output; existing files are never overwritten." }
        $parent = [IO.Path]::GetDirectoryName($OutputPath)
        if (-not [IO.Directory]::Exists($parent)) { throw "PDF output directory does not exist: '$parent'." }
        if ($null -eq $Staging) {
            $ownedStaging = New-PdfStaging -OutputFolder $parent
            $Staging = $ownedStaging
        }
        Assert-PdfStaging -Staging $Staging
        if ($parent -ine $Staging.OutputFolder) { throw 'Private staging must belong to this output directory.' }
        $stagedOutput = if ($Tool -eq 'Pdftk') { $Staging.MasterPath } else { $Staging.EmailPath }
        if (Test-Path -LiteralPath $stagedOutput) { throw 'Owned staged PDF already exists. Refusing native overwrite; use a fresh run.' }
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
        $validationLabel = if ($Tool -eq 'Pdftk') { 'Master' } else { 'Email' }
        $inspector = if ($Tool -eq 'Pdftk') { $Executable } else { $InspectionExecutable }
        # Native success is only staging. Inspect every owned PDF before the
        # no-overwrite move can create a final master or smaller email copy.
        Assert-PdfStaging -Staging $Staging
        $stagedFile = Get-Item -LiteralPath $stagedOutput -Force -ErrorAction Stop
        if ($stagedFile -isnot [IO.FileInfo] -or
            ($stagedFile.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "$validationLabel validation requires a regular owned staged PDF; no final output was published."
        }
        $snapshot = Get-PdfInputSnapshot -LiteralPath $stagedOutput
        $validation = Get-PdfDocumentInspection -Executable $inspector -LiteralPath $stagedOutput -TimeoutMilliseconds $TimeoutMilliseconds
        if (-not $validation.Succeeded -or $null -eq $validation.NativeResult -or -not $validation.NativeResult.Succeeded) {
            throw "$validationLabel validation failed; no final output was published. $($validation.InputError)"
        }
        $validationNative = $validation.NativeResult
        foreach ($field in @('Started','ExitCode','Succeeded','TimedOut','Cancelled','LaunchError','CaptureError','TerminationError','StdoutTruncated','StderrTruncated')) {
            if ($null -eq $validationNative.PSObject.Properties[$field]) {
                throw "$validationLabel validation receipt is incomplete ($field); no final output was published."
            }
        }
        if (-not $validationNative.Started -or $validationNative.ExitCode -ne 0 -or
            $validationNative.TimedOut -or $validationNative.Cancelled -or $validationNative.LaunchError -or
            $validationNative.CaptureError -or $validationNative.TerminationError) {
            throw "$validationLabel validation native execution was incomplete or unsuccessful; no final output was published."
        }
        if ($validation.NativeResult.StdoutTruncated -or $validation.NativeResult.StderrTruncated) {
            throw "$validationLabel validation capture was incomplete; no final output was published."
        }
        if ($validation.PageCount -ne $ExpectedPageCount) {
            throw "$validationLabel validation page count mismatch: expected $ExpectedPageCount, inspected $($validation.PageCount). No final output was published."
        }
        Assert-PdfStaging -Staging $Staging
        $current = Get-PdfInputSnapshot -LiteralPath $stagedOutput
        $currentFile = Get-Item -LiteralPath $stagedOutput -Force -ErrorAction Stop
        if (($currentFile.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "Staged $validationLabel became a reparse file during validation; no final output was published."
        }
        if ($current.Length -ne $snapshot.Length -or $current.LastWriteTimeUtcTicks -ne $snapshot.LastWriteTimeUtcTicks) {
            throw "Staged $validationLabel changed during validation; no final output was published. Use stable source documents and run again."
        }
        if ($Tool -eq 'Ghostscript') {
            $currentMaster = Get-PdfInputSnapshot -LiteralPath $InputPaths[0]
            $currentMasterFile = Get-Item -LiteralPath $InputPaths[0] -Force -ErrorAction Stop
            if (($currentMasterFile.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0 -or
                $currentMaster.Length -ne $masterSnapshot.Length -or
                $currentMaster.LastWriteTimeUtcTicks -ne $masterSnapshot.LastWriteTimeUtcTicks) {
                throw 'Published master changed during email processing; no email output was published. Use stable documents and run again.'
            }
        }
        $validatedPages = [long]$validation.PageCount
        $outputBytes = $current.Length
        $validated = $true
        # File.Move refuses an existing target, including a collision after
        # validation and the strictly smaller email decision.
        if ($Tool -eq 'Ghostscript' -and $outputBytes -ge $masterBytes) {
            $outputState = 'no_size_benefit'
        } else {
            Publish-PdfStagedOutput -Staging $Staging -StagedPath $stagedOutput -OutputPath $OutputPath
            $published = $true
            $outputState = 'published'
        }
    } catch {
        $outputError = $_.Exception.Message
    } finally {
        if ($null -ne $ownedStaging) {
            $cleanup = Remove-PdfStaging -Staging $ownedStaging
            $cleanupError = $cleanup.CleanupError
        }
    }
    return [pscustomobject]@{
        NativeResult = $native
        ValidationResult = $validation
        OutputValidated = $validated
        ValidatedPageCount = $validatedPages
        OutputState = $outputState
        MasterBytes = $masterBytes
        OutputBytes = $outputBytes
        OutputPath = $OutputPath
        OutputPublished = $published
        OutputError = $outputError
        CleanupError = $cleanupError
        StagingPath = $(if ($null -ne $Staging) { $Staging.DirectoryPath } else { $null })
        Succeeded = ($validated -and -not $outputError -and ($published -or $outputState -eq 'no_size_benefit'))
    }
}

function Get-PdfMergeOutcome {
    param(
        [Parameter(Mandatory=$true)][bool]$MasterPublished,
        [ValidateSet('not_started','skipped','unavailable','published','no_size_benefit','failed')][string]$EmailState = 'not_started',
        [string]$MasterPath,
        [string]$EmailPath
    )

    $paths = @()
    $exitCode = 1
    $summary = 'FAILURE'
    $message = 'No validated master was published.'
    if ($MasterPublished) {
        if ([string]::IsNullOrWhiteSpace($MasterPath)) { throw 'A published master requires its explicit output path.' }
        $paths += [pscustomobject]@{ Label='Merged master'; Path=$MasterPath }
        $exitCode = 0
        $summary = 'SUCCESS'
        switch ($EmailState) {
            'skipped' { $message = 'Email explicitly skipped; validated master retained.' }
            'unavailable' { $message = 'Ghostscript not found; skipping email-optimized copy.' }
            'no_size_benefit' { $message = 'Validated email copy offers no size benefit; master retained.' }
            'published' {
                if ([string]::IsNullOrWhiteSpace($EmailPath)) { throw 'A published email result requires its explicit output path.' }
                $message = 'Validated smaller email copy published.'
                $paths += [pscustomobject]@{ Label='Email-optimized'; Path=$EmailPath }
            }
            default {
                $exitCode = 2
                $summary = 'PARTIAL SUCCESS'
                $message = 'Email processing failed; validated master retained.'
            }
        }
    }
    return [pscustomobject]@{ ExitCode=$exitCode; Summary=$summary; EmailMessage=$message; EmailState=$EmailState; PublishedPaths=@($paths) }
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
