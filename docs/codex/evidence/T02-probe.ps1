# Development-only T02 observation procedure. Never imports the application entry.
# Runs isolated baseline functions and synthetic shell/process probes, not a PDF merge.
param([Parameter(Mandatory=$true)][string]$PythonPath)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path
$entry = Join-Path $repo 'WinPDFMerge.ps1'
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($entry, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -ne 0) { throw 'Baseline parser errors; inspect before probing.' }
# Import only this AST-selected function definition, never the entry orchestration.
$sortFunction = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'NaturalSortKey' }, $true)
. ([scriptblock]::Create($sortFunction.Extent.Text))
$work = Join-Path ([IO.Path]::GetTempPath()) ('WinPDFMerger-T02-' + [Guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($work)
$observations = [ordered]@{}
try {
    $os = Get-CimInstance Win32_OperatingSystem
    $windows = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
    $policies = @(Get-ExecutionPolicy -List | ForEach-Object { [ordered]@{scope=$_.Scope.ToString(); policy=$_.ExecutionPolicy.ToString()} })
    $modules = @(Get-Module -ListAvailable Pester,PSScriptAnalyzer | ForEach-Object { [ordered]@{name=$_.Name; version=$_.Version.ToString()} })
    $volume = Get-Volume -DriveLetter ([IO.Path]::GetPathRoot($repo).Substring(0,1))
    $principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $observations.environment = [ordered]@{
        observed_at_utc=[DateTime]::UtcNow.ToString('o'); commit_under_test=(& git -C $repo rev-parse HEAD)
        os_caption=$os.Caption; os_version=$os.Version; os_build=$os.BuildNumber; os_architecture=$os.OSArchitecture
        display_version=$windows.DisplayVersion; build_lab=$windows.BuildLabEx; installation_type=$windows.InstallationType
        shell_version=$PSVersionTable.PSVersion.ToString(); shell_edition=$PSVersionTable.PSEdition
        process_64_bit=[Environment]::Is64BitProcess; os_64_bit=[Environment]::Is64BitOperatingSystem
        token_is_administrator=$principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        effective_execution_policy=(Get-ExecutionPolicy).ToString(); execution_policies=$policies; available_test_modules=$modules
        repository_filesystem=$volume.FileSystem; repository_drive_type=$volume.DriveType.ToString()
        python_version=(& $PythonPath --version); git_version=(& git --version); gh_version=(& gh --version | Select-Object -First 1)
    }
    $observations.probe_sha256 = (Get-FileHash -LiteralPath $PSCommandPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $observations.parser_error_count = $parseErrors.Count
    $observations.source_hashes = @(foreach ($name in @('LICENSE','README.md','WinPDFMerge.bat','WinPDFMerge.ico','WinPDFMerge.ps1','WinPDFMerge_icon.png','WinPDFMerge_poster.png')) {
        $hash = Get-FileHash -LiteralPath (Join-Path $repo $name) -Algorithm SHA256
        [ordered]@{file=$name; sha256=$hash.Hash.ToLowerInvariant()}
    })
    $observations.dependency_commands = @(foreach ($name in @('pdftk','pdftk.exe','gswin64c.exe','gswin32c.exe')) {
        [ordered]@{name=$name; command_found=($null -ne (Get-Command $name -ErrorAction SilentlyContinue))}
    })
    $locations = @()
    foreach ($root in @($env:ProgramFiles, ${env:ProgramFiles(x86)})) {
        foreach ($relative in @('PDFtk Server\bin\pdftk.exe','PDFtk\bin\pdftk.exe')) {
            $locations += [ordered]@{path=(Join-Path $root $relative); exists=(Test-Path -LiteralPath (Join-Path $root $relative) -PathType Leaf)}
        }
        $gsRoot = Join-Path $root 'gs'
        $locations += [ordered]@{path=$gsRoot; exists=(Test-Path -LiteralPath $gsRoot -PathType Container)}
        if (Test-Path -LiteralPath $gsRoot) {
            foreach ($folder in Get-ChildItem -LiteralPath $gsRoot -Directory) {
                foreach ($exe in @('gswin64c.exe','gswin32c.exe')) {
                    $candidate = Join-Path $folder.FullName ('bin\' + $exe)
                    $locations += [ordered]@{path=$candidate; exists=(Test-Path -LiteralPath $candidate -PathType Leaf)}
                }
            }
        }
    }
    $observations.dependency_locations = $locations
    $observations.dependency_registrations = @(foreach ($key in @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*','HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*')) {
        Get-ItemProperty $key -ErrorAction SilentlyContinue | Where-Object { $_.PSObject.Properties['DisplayName'] -and $_.DisplayName -match 'PDFtk|Ghostscript' } | ForEach-Object { [ordered]@{name=$_.DisplayName; version=$_.DisplayVersion} }
    })
    try { $null = NaturalSortKey '2147483648'; $observations.int32_overflow = 'not observed' }
    catch { $observations.int32_overflow = $_.Exception.GetType().Name }
    $observations.sort_orders = @(foreach ($cultureName in @('en-US','de-DE')) {
        $savedCulture = [Threading.Thread]::CurrentThread.CurrentCulture
        try {
            [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo($cultureName)
            $items = @('10','2','01','1') | ForEach-Object { [pscustomobject]@{BaseName=$_; FullName=('C:\Synthetic\' + $_ + '.pdf')} }
            [ordered]@{culture=$cultureName; input=@('10','2','01','1'); order=@($items | Sort-Object { NaturalSortKey $_.BaseName }, FullName | ForEach-Object {$_.BaseName})}
        } finally { [Threading.Thread]::CurrentThread.CurrentCulture = $savedCulture }
    })
    $observations.fallback_x86 = [ordered]@{baseline="$Env:ProgramFiles(x86)\PDFtk\bin\pdftk.exe"; braced="${env:ProgramFiles(x86)}\PDFtk\bin\pdftk.exe"}
    $source = Join-Path $work 'Source[1]'
    [void][IO.Directory]::CreateDirectory($source)
    [IO.File]::WriteAllText((Join-Path $source '1.PDF'), 'synthetic non-PDF marker; enumeration only')
    $observations.bracket_path = [ordered]@{
        resolve_path_matches=@(Resolve-Path -Path $source -ErrorAction SilentlyContinue).Count
        resolve_literal_matches=@(Resolve-Path -LiteralPath $source).Count
        test_path=(Test-Path -Path $source -PathType Container); test_literal_path=(Test-Path -LiteralPath $source -PathType Container)
    }
    $pdfs = Get-ChildItem -LiteralPath $source -Filter *.pdf -File -ErrorAction Stop
    try { $observations.singleton_count = $pdfs.Count }
    catch { $observations.singleton_count = $_.Exception.GetType().Name }
    $observations.uppercase_enumeration_count = @($pdfs).Count
    $observations.root_leaf = Split-Path 'C:\' -Leaf
    # A real Windows process, using Python ONLY as an argument observer. No PDF engine.
    $echoFile = Join-Path $work 'echo_args.py'
    [IO.File]::WriteAllText($echoFile, 'import json,sys; print(json.dumps(sys.argv[1:]))', [Text.Encoding]::ASCII)
    $out = Join-Path $work 'echo.stdout.txt'
    $err = Join-Path $work 'echo.stderr.txt'
    $echoArgs = @(('"{0}"' -f $echoFile),'"C:\Synthetic Input\1.pdf"','cat','output','C:\Synthetic Output\master.pdf','compress')
    $process = Start-Process -FilePath $PythonPath -ArgumentList $echoArgs -WindowStyle Hidden -PassThru -RedirectStandardOutput $out -RedirectStandardError $err
    if (-not $process.WaitForExit(10000)) { $process.Kill(); throw 'Argument observer timed out.' }
    if ($process.ExitCode -ne 0) { throw 'Argument observer failed.' }
    $observations.start_process_received_arguments = @(Get-Content -LiteralPath $out -Raw | ConvertFrom-Json)
    # The baseline's delayed expansion / unquoted assignment primitives only.
    # Do not launch its ExecutionPolicy Bypass command or claim Explorer acceptance.
    $batchFile = Join-Path $work 'batch_probe.bat'
    [IO.File]::WriteAllLines($batchFile, @('@echo off','setlocal ENABLEDELAYEDEXPANSION','set "WINPDFMERGER_T02_SYNTHETIC="','set SCRIPT_DIR=C:\Synthetic!WINPDFMERGER_T02_SYNTHETIC!\','echo DELAYED=[%SCRIPT_DIR%]','set SCRIPT_DIR=C:\Synthetic&echo AMPERSAND_COMMAND_EXECUTED','echo ASSIGNMENT=[%SCRIPT_DIR%]'), [Text.Encoding]::ASCII)
    $observations.batch_primitive_output = @(& $env:ComSpec /d /c $batchFile)
    $observations.gs_lexical_directory_choice = @('gs9.56.1','gs10.06.0') | Sort-Object -Descending | Select-Object -First 1
    $observations.pdf_merge_executed = $false
    $observations.explorer_drag_drop_executed = $false
    $observations | ConvertTo-Json -Depth 8
} finally {
    # Only this invocation's verified direct child of TEMP is removed.
    $resolvedWork = [IO.Path]::GetFullPath($work)
    $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if (-not $resolvedWork.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -or [IO.Path]::GetFileName($resolvedWork) -notmatch '^WinPDFMerger-T02-[0-9a-f]{32}$') {
        throw 'Refusing cleanup outside the owned T02 directory.'
    }
    Remove-Item -LiteralPath $resolvedWork -Recurse -Force
}
