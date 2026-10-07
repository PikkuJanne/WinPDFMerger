# Internal baseline helpers. Import defines functions only; entry orchestration stays in WinPDFMerge.ps1.
# Keep the measured behavior here until the corresponding regression task changes it.

function Find-Pdftk {
    $pdftk = Get-Command pdftk -ErrorAction SilentlyContinue
    if ($pdftk) { return $pdftk.Source }
    $candidates = @(
        "$Env:ProgramFiles\PDFtk Server\bin\pdftk.exe",
        "$Env:ProgramFiles(x86)\PDFtk\bin\pdftk.exe",
        "$Env:ProgramFiles\Pdftk Server\bin\pdftk.exe"
    )
    foreach ($c in $candidates) { if (Test-Path $c) { return $c } }
    return $null
}
function Find-Ghostscript {
    $gs = Get-Command gswin64c.exe -ErrorAction SilentlyContinue
    if ($gs) { return $gs.Source }
    $gs = Get-Command gswin32c.exe -ErrorAction SilentlyContinue
    if ($gs) { return $gs.Source }
    $common = Get-ChildItem -Path "$Env:ProgramFiles\gs" -Directory -ErrorAction SilentlyContinue |
              Sort-Object Name -Descending | Select-Object -First 1
    if ($common) {
        $cand = Join-Path $common.FullName "bin\gswin64c.exe"
        if (Test-Path $cand) { return $cand }
    }
    return $null
}
function NaturalSortKey([string]$s) {
    [regex]::Split($s, '(\d+)') | ForEach-Object { if ($_ -match '^\d+$') { [int]$_ } else { $_ } }
}
function Sanitize-FileName([string]$name) {
    $invalid = [IO.Path]::GetInvalidFileNameChars() -join ''
    $re = "[{0}]" -f ([Regex]::Escape($invalid))
    ($name -replace $re, '_').Trim()
}
