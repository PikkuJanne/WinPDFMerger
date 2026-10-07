BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
    . (Join-Path $repo 'src/WinPDFMerge.Helpers.ps1')
    function New-ControlledPdfEnvelope([string]$LineEnding = "`n", [switch]$IndirectTarget) {
        # These are byte-envelope unit shapes, not native-valid PDF fixtures.
        $prefix = '%PDF-1.5' + $LineEnding + '1 0 obj << /Controlled true >> endobj' + $LineEnding
        $offset = [Text.Encoding]::GetEncoding(28591).GetByteCount($prefix)
        $target = if ($IndirectTarget) { '2 0 obj << /Type /XRef >>' } else { 'xref' }
        $body = $prefix + $target + $LineEnding + 'startxref' + $LineEnding + $offset.ToString([Globalization.CultureInfo]::InvariantCulture) + $LineEnding + '%%EOF' + $LineEnding
        [pscustomobject]@{ Text=$body; Offset=$offset }
    }
}

Describe 'T11: bounded PDF byte-envelope plausibility' {
    BeforeEach { $path = Join-Path $TestDrive ('envelope-' + [Guid]::NewGuid().ToString('N') + '.pdf') }

    It 'accepts a conventional target with <Label> and preserves exact bytes' -TestCases @(
        @{Label='LF'; Ending="`n"}, @{Label='CR'; Ending="`r"}, @{Label='CRLF'; Ending="`r`n"}
    ) {
        param($Label,$Ending)
        $shape = New-ControlledPdfEnvelope -LineEnding $Ending
        [IO.File]::WriteAllBytes($path,[Text.Encoding]::GetEncoding(28591).GetBytes($shape.Text))
        $before = (Get-FileHash -LiteralPath $path).Hash
        { Assert-PdfInputEnvelope -LiteralPath $path } | Should -Not -Throw
        (Get-FileHash -LiteralPath $path).Hash | Should -BeExactly $before
    }

    It 'accepts an indirect-object target shape for cross-reference streams' {
        $shape = New-ControlledPdfEnvelope -IndirectTarget
        [IO.File]::WriteAllBytes($path,[Text.Encoding]::ASCII.GetBytes($shape.Text))
        { Assert-PdfInputEnvelope -LiteralPath $path } | Should -Not -Throw
    }

    It 'accepts a PDF comment as the delimiter immediately after obj' {
        $shape = New-ControlledPdfEnvelope -IndirectTarget
        [IO.File]::WriteAllText($path,$shape.Text.Replace('2 0 obj <<',"2 0 obj%controlled comment`n<<"),[Text.Encoding]::ASCII)
        { Assert-PdfInputEnvelope -LiteralPath $path } | Should -Not -Throw
    }

    It 'accepts comments separating indirect-object header tokens' {
        $shape = New-ControlledPdfEnvelope -IndirectTarget
        [IO.File]::WriteAllText($path,$shape.Text.Replace('2 0 obj <<',"2%number comment`r0%generation comment`nobj <<"),[Text.Encoding]::ASCII)
        { Assert-PdfInputEnvelope -LiteralPath $path } | Should -Not -Throw
    }

    It 'rejects a case-changed <Label> PDF keyword' -TestCases @(
        @{Label='header'; From='%PDF'; To='%pdf'; Indirect=$false},
        @{Label='xref'; From='xref'; To='XREF'; Indirect=$false},
        @{Label='obj'; From='obj'; To='OBJ'; Indirect=$true},
        @{Label='footer'; From='startxref'; To='STARTXREF'; Indirect=$false}
    ) {
        param($Label,$From,$To,$Indirect)
        $shape = New-ControlledPdfEnvelope -IndirectTarget:$Indirect
        $text = if ($Label -eq 'xref') { $shape.Text.Replace("`nxref`n","`nXREF`n") } else { $shape.Text.Replace($From,$To) }
        [IO.File]::WriteAllText($path,$text,[Text.Encoding]::ASCII)
        { Assert-PdfInputEnvelope -LiteralPath $path } | Should -Throw '*envelope*'
    }

    It 'accepts horizontal whitespace around the final numeric offset' {
        $shape = New-ControlledPdfEnvelope
        $text = $shape.Text.Replace(("startxref`n" + $shape.Offset), ("startxref`n `t" + $shape.Offset + " `t"))
        [IO.File]::WriteAllText($path,$text,[Text.Encoding]::ASCII)
        { Assert-PdfInputEnvelope -LiteralPath $path } | Should -Not -Throw
    }

    It 'uses byte offsets when the prefix contains binary characters' {
        $shape = New-ControlledPdfEnvelope
        $binary = [char]0xe2 + [string][char]0xe3 + [char]0xcf + [char]0xd3
        $prefix = "%PDF-1.5`n%" + $binary + "`n"
        $offset = $prefix.Length
        $text = $prefix + "xref`nstartxref`n" + $offset + "`n%%EOF`n"
        [IO.File]::WriteAllBytes($path,[Text.Encoding]::GetEncoding(28591).GetBytes($text))
        { Assert-PdfInputEnvelope -LiteralPath $path } | Should -Not -Throw
    }

    It 'uses the final footer after an earlier dummy startxref0 and EOF' {
        $prefix = "%PDF-1.4`nstartxref`n0`n%%EOF`n"
        $offset = $prefix.Length
        [IO.File]::WriteAllText($path,($prefix + "xref`nstartxref`n" + $offset + "`n%%EOF`n"),[Text.Encoding]::ASCII)
        { Assert-PdfInputEnvelope -LiteralPath $path } | Should -Not -Throw
    }

    It 'allows PDF whitespace after the terminal EOF' {
        $shape = New-ControlledPdfEnvelope
        [IO.File]::WriteAllText($path,($shape.Text + " `t`r`n" + [char]0 + [char]12),[Text.Encoding]::ASCII)
        { Assert-PdfInputEnvelope -LiteralPath $path } | Should -Not -Throw
    }

    It 'rejects <Label> instead of trusting a backend-repaired count' -TestCases @(
        @{Label='missing header'; Mode='header'},
        @{Label='missing EOF'; Mode='eof'},
        @{Label='zero final offset'; Mode='zero'},
        @{Label='out-of-file offset'; Mode='outside'},
        @{Label='overflow offset'; Mode='overflow'},
        @{Label='invalid target'; Mode='target'},
        @{Label='trailing garbage'; Mode='garbage'},
        @{Label='footer beyond supported window'; Mode='window'}
    ) {
        param($Label,$Mode)
        $shape = New-ControlledPdfEnvelope
        $text = switch ($Mode) {
            'header' { $shape.Text.Replace('%PDF-1.5','NOTPDF!!') }
            'eof' { $shape.Text.Replace('%%EOF','') }
            'zero' { $shape.Text.Replace(('startxref' + "`n" + $shape.Offset),"startxref`n0") }
            'outside' { $shape.Text.Replace(('startxref' + "`n" + $shape.Offset),"startxref`n999999") }
            'overflow' { $shape.Text.Replace(('startxref' + "`n" + $shape.Offset),"startxref`n9223372036854775808") }
            'target' { $shape.Text.Replace("`nxref`n","`nxr-f`n") }
            'garbage' { $shape.Text + 'not PDF whitespace' }
            'window' { $shape.Text + (' ' * 8192) }
        }
        [IO.File]::WriteAllText($path,$text,[Text.Encoding]::ASCII)
        { Assert-PdfInputEnvelope -LiteralPath $path } | Should -Throw '*envelope*'
        # The read-only handle is closed on rejection as well as success.
        $handle = [IO.File]::Open($path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::None)
        $handle.Dispose()
    }
}
