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
        # Tee-Object's LiteralPath parameter set has no Append in either shell.
        # Out-File keeps the existing per-shell encoding and console echo.
        $Message | Out-File -LiteralPath $LiteralPath -Append:$writeAppend -ErrorAction Stop
        $writeAppend = $true
        $Message
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

    # This bounded, fixed-argument preflight is separate from the general native
    # argument/lifecycle work. Start both reads before waiting for the owned child.
    $startInfo = New-Object Diagnostics.ProcessStartInfo
    $startInfo.FileName = $Path
    $startInfo.Arguments = '--version'
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardInput = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.EnvironmentVariables.Remove('GS_OPTIONS')
    $process = New-Object Diagnostics.Process
    $process.StartInfo = $startInfo
    $started = $false
    $timer = [Diagnostics.Stopwatch]::StartNew()
    try {
        $started = $process.Start()
        if (-not $started) { throw 'Version probe did not start.' }
        $process.StandardInput.Close()
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        $remaining = [Math]::Max(0, $TimeoutMilliseconds - [int]$timer.ElapsedMilliseconds)
        if (-not $process.WaitForExit($remaining)) { throw "Version probe timed out after $TimeoutMilliseconds ms." }
        $remaining = [Math]::Max(0, $TimeoutMilliseconds - [int]$timer.ElapsedMilliseconds)
        if (-not [Threading.Tasks.Task]::WaitAll([Threading.Tasks.Task[]]@($stdout, $stderr), $remaining)) {
            throw "Version probe stream capture timed out after $TimeoutMilliseconds ms."
        }
        return [pscustomobject]@{ ExitCode = $process.ExitCode; Stdout = $stdout.Result; Stderr = $stderr.Result }
    } finally {
        # Terminate only this probe, never other PDFtk/GS sessions by image name.
        # Descendant cancellation and general native cleanup remain later work.
        if ($started -and -not $process.HasExited) {
            try {
                $process.Kill()
                if (-not $process.WaitForExit(1000)) { Write-Warning 'Version probe termination was best effort.' }
            } catch {
                Write-Warning ('Version probe termination was best effort: ' + $_.Exception.Message)
            }
        }
        $timer.Stop()
        $process.Dispose()
    }
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
