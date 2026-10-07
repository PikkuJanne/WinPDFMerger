# Development-only. Compile the controlled process fixture; never installs tools.
# All generated files stay in one unique directory under tests/.work.
[CmdletBinding()]
param([string]$CompilerPath)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) {
    throw 'FakeNative must be built and tested on Windows.'
}
$repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
$source = Join-Path $repo 'tests/native/FakeNative.cs'
if (-not (Test-Path -LiteralPath $source -PathType Leaf)) {
    throw 'Missing controlled-process source tests/native/FakeNative.cs.'
}

if (-not $CompilerPath) {
    foreach ($relative in @('Microsoft.NET/Framework64/v4.0.30319/csc.exe', 'Microsoft.NET/Framework/v4.0.30319/csc.exe')) {
        $candidate = Join-Path $env:WINDIR $relative
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            $CompilerPath = $candidate
            break
        }
    }
}
if (-not $CompilerPath -or -not (Test-Path -LiteralPath $CompilerPath -PathType Leaf)) {
    throw 'No existing Windows C# compiler found. Supply -CompilerPath to an existing csc.exe; this helper does not install dependencies.'
}
$compiler = (Resolve-Path -LiteralPath $CompilerPath).ProviderPath
if ([IO.Path]::GetFileName($compiler) -ine 'csc.exe') {
    throw 'CompilerPath must identify csc.exe.'
}
$workRoot = Join-Path $repo 'tests/.work/fake-native'
$work = Join-Path $workRoot ([Guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($work)
$executable = Join-Path $work 'FakeNative.exe'
$compilerOutput = @(& $compiler '/nologo' '/target:exe' '/platform:anycpu' '/optimize+' ("/out:{0}" -f $executable) $source 2>&1)
$compilerExitCode = $LASTEXITCODE
$compilerLog = Join-Path $work 'compiler.txt'
[IO.File]::WriteAllLines($compilerLog, [string[]]$compilerOutput, [Text.UTF8Encoding]::new($false))
if ($compilerExitCode -ne 0 -or -not (Test-Path -LiteralPath $executable -PathType Leaf)) {
    throw "FakeNative build failed (exit $compilerExitCode). Inspect $compilerLog."
}

$receipt = [ordered]@{
    purpose = 'Controlled argument/fault process; not PDF/native-engine evidence'
    compiler = $compiler
    compiler_version = (Get-Item -LiteralPath $compiler).VersionInfo.FileVersion
    source_sha256 = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash.ToLowerInvariant()
    executable_sha256 = (Get-FileHash -LiteralPath $executable -Algorithm SHA256).Hash.ToLowerInvariant()
    shell_version = $PSVersionTable.PSVersion.ToString()
    built_at_utc = [DateTime]::UtcNow.ToString('o')
}
[IO.File]::WriteAllText((Join-Path $work 'build-info.json'), ($receipt | ConvertTo-Json -Depth 3), [Text.UTF8Encoding]::new($false))
# A single path on the success stream keeps the caller's process setup explicit.
$executable
