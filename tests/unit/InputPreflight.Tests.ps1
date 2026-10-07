BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
    . (Join-Path $repo 'src/WinPDFMerge.Helpers.ps1')
    function New-InputPreflightResult([string]$Text = "NumberOfPages: 2`n") {
        [pscustomobject]@{
            Succeeded = $true; Started = $true; ExitCode = 0; TimedOut = $false
            Cancelled = $false; LaunchError = $null; CaptureError = $null; TerminationError = $null
            StdoutTruncated = $false; StderrTruncated = $false; Stdout = $Text; Stderr = ''
        }
    }
}

Describe 'T11: document-data page count parsing' {
    It 'uses the one complete labeled count amid unrelated metadata and noise' {
        $text = "InfoValue: NumberOfPages: 99999`r`nWarning: 42 pages mentioned`r`nNumberOfPages: `t0002 `t`r`nPageMediaNumber: 1`r`n"
        ConvertFrom-PdfDocumentData -Text $text | Should -Be 2
    }

    It 'parses a positive Int64 count independently of culture' {
        $saved = [Threading.Thread]::CurrentThread.CurrentCulture
        try {
            [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]'ar-SA'
            ConvertFrom-PdfDocumentData -Text 'NumberOfPages: 2147483648' | Should -Be ([long]2147483648)
        } finally { [Threading.Thread]::CurrentThread.CurrentCulture = $saved }
    }

    It 'rejects <Label> without inventing a count' -TestCases @(
        @{ Label='empty output'; Text='' },
        @{ Label='digits without a label'; Text='Warning: 77 pages, file version 1.7' },
        @{ Label='count inside metadata only'; Text='InfoValue: NumberOfPages: 6' },
        @{ Label='split label/value'; Text="NumberOfPages:`n2" },
        @{ Label='zero pages'; Text='NumberOfPages: 0' },
        @{ Label='negative count'; Text='NumberOfPages: -2' },
        @{ Label='explicit plus'; Text='NumberOfPages: +2' },
        @{ Label='fraction'; Text='NumberOfPages: 2.0' },
        @{ Label='exponent'; Text='NumberOfPages: 2e3' },
        @{ Label='non-ASCII digits'; Text=('NumberOfPages: ' + [char]0x0662) },
        @{ Label='trailing text'; Text='NumberOfPages: 2 warning' },
        @{ Label='overflow'; Text='NumberOfPages: 9223372036854775808' },
        @{ Label='duplicate label'; Text="NumberOfPages: 2`nNumberOfPages: 2" },
        @{ Label='malformed duplicate'; Text="NumberOfPages: 2`nNumberOfPages: broken" }
    ) {
        param($Label,$Text)
        { ConvertFrom-PdfDocumentData -Text $Text } | Should -Throw '*page count*'
    }
}

Describe 'T11: frozen input metadata' {
    BeforeEach {
        $path = Join-Path $TestDrive ('source [x] & ! ' + [Guid]::NewGuid().ToString('N') + '.pdf')
        [IO.File]::WriteAllText($path,'synthetic controlled source')
    }

    It 'records literal path, length and UTC timestamp without editing the source' {
        $before = (Get-FileHash -LiteralPath $path).Hash
        $snapshot = Get-PdfInputSnapshot -LiteralPath $path
        $snapshot.FullName | Should -BeExactly $path
        $snapshot.Length | Should -Be ([IO.FileInfo]$path).Length
        $snapshot.LastWriteTimeUtcTicks | Should -Be ([IO.FileInfo]$path).LastWriteTimeUtc.Ticks
        { Assert-PdfInputSnapshot -Snapshot $snapshot } | Should -Not -Throw
        (Get-FileHash -LiteralPath $path).Hash | Should -BeExactly $before
    }

    It 'rejects empty input by name' {
        [IO.File]::WriteAllBytes($path,[byte[]]@())
        { Get-PdfInputSnapshot -LiteralPath $path } | Should -Throw '*empty*'
    }

    It 'rejects missing input by name' {
        $missing = Join-Path $TestDrive 'missing.pdf'
        { Get-PdfInputSnapshot -LiteralPath $missing } | Should -Throw '*missing.pdf*'
    }

    It 'rejects a directory instead of an input file' {
        { Get-PdfInputSnapshot -LiteralPath $TestDrive } | Should -Throw '*file*'
    }

    It 'detects a length change after the snapshot' {
        $snapshot = Get-PdfInputSnapshot -LiteralPath $path
        [IO.File]::AppendAllText($path,'changed')
        { Assert-PdfInputSnapshot -Snapshot $snapshot } | Should -Throw '*changed*'
    }

    It 'detects timestamp change even when length is unchanged' {
        $snapshot = Get-PdfInputSnapshot -LiteralPath $path
        [IO.File]::SetLastWriteTimeUtc($path,([IO.FileInfo]$path).LastWriteTimeUtc.AddSeconds(2))
        { Assert-PdfInputSnapshot -Snapshot $snapshot } | Should -Throw '*changed*'
    }
}

Describe 'T11: bounded read-only PDFtk inspection wiring' {
    BeforeEach {
        $path = Join-Path $TestDrive ('input [x] & ! ' + [Guid]::NewGuid().ToString('N') + '.pdf')
        [IO.File]::WriteAllText($path,'controlled input, native call mocked')
        $exe = Join-Path $TestDrive 'pdftk.exe'
        [IO.File]::WriteAllText($exe,'controlled placeholder, never executed')
        Mock Assert-PdfInputEnvelope { }
    }

    It 'uses one absolute input, fixed document-data operation and dont_ask with injected bound' {
        Mock Invoke-NativeProcess {
            param($Executable,$Arguments,$TimeoutMilliseconds)
            $Executable | Should -BeExactly $exe
            @($Arguments).Count | Should -Be 5
            $Arguments[0] | Should -BeExactly $path
            ($Arguments[1..4] -join '|') | Should -BeExactly 'dump_data_utf8|output|-|dont_ask'
            $TimeoutMilliseconds | Should -Be 3210
            New-InputPreflightResult
        }
        $result = Get-PdfDocumentInspection -Executable $exe -LiteralPath $path -TimeoutMilliseconds 3210
        $result.Succeeded | Should -BeTrue -Because $result.InputError
        $result.PageCount | Should -Be 2
        $result.NativeResult | Should -Not -BeNullOrEmpty
        @(Get-ChildItem -LiteralPath $TestDrive -File).Count | Should -Be 2
        Should -Invoke Invoke-NativeProcess -Times 1 -Exactly
    }

    It 'rejects <Label> even when stdout contains a plausible count' -TestCases @(
        @{ Label='nonzero exit'; Field='ExitCode'; Value=1 },
        @{ Label='timeout'; Field='TimedOut'; Value=$true },
        @{ Label='cancellation'; Field='Cancelled'; Value=$true },
        @{ Label='capture failure'; Field='CaptureError'; Value='controlled capture failure' },
        @{ Label='truncated stdout'; Field='StdoutTruncated'; Value=$true },
        @{ Label='termination failure'; Field='TerminationError'; Value='controlled termination failure' }
    ) {
        param($Label,$Field,$Value)
        Mock Invoke-NativeProcess {
            $native = New-InputPreflightResult
            $native.$Field = $Value
            $native.Succeeded = $false
            $native
        }
        $result = Get-PdfDocumentInspection -Executable $exe -LiteralPath $path
        $result.Succeeded | Should -BeFalse
        $result.PageCount | Should -BeNullOrEmpty
        $result.InputError | Should -Match ([regex]::Escape($path))
    }

    It 'never treats a count in stderr as document data' {
        Mock Invoke-NativeProcess {
            $native = New-InputPreflightResult -Text 'Warning: no count in document data'
            $native.Stderr = 'NumberOfPages: 999'
            $native
        }
        $result = Get-PdfDocumentInspection -Executable $exe -LiteralPath $path
        $result.Succeeded | Should -BeFalse
        $result.InputError | Should -Match 'page count'
    }

    It 'refuses an unsupported long operand before launching the native process' {
        Mock Invoke-NativeProcess { throw 'Unexpected native launch' }
        $longPath = 'C:\' + ('a' * 253) + '.pdf'
        $result = Get-PdfDocumentInspection -Executable $exe -LiteralPath $longPath
        $result.Succeeded | Should -BeFalse
        $result.InputError | Should -Match '260'
        Should -Invoke Invoke-NativeProcess -Times 0 -Exactly
    }

    It 'retains benign stderr warnings while using only the labeled stdout count' {
        Mock Invoke-NativeProcess {
            $native = New-InputPreflightResult
            $native.Stderr = 'Controlled warning with 999 unrelated digits'
            $native
        }
        $result = Get-PdfDocumentInspection -Executable $exe -LiteralPath $path
        $result.Succeeded | Should -BeTrue
        $result.PageCount | Should -Be 2
        $result.NativeResult.Stderr | Should -BeExactly 'Controlled warning with 999 unrelated digits'
    }
}

Describe 'T11: complete ordered page inventory' {
    BeforeEach {
        $one = Join-Path $TestDrive ('one-' + [Guid]::NewGuid().ToString('N') + '.pdf')
        $two = Join-Path $TestDrive ('two-' + [Guid]::NewGuid().ToString('N') + '.pdf')
        [IO.File]::WriteAllText($one,'one controlled input')
        [IO.File]::WriteAllText($two,'two controlled input')
        $inputs = @((Get-Item -LiteralPath $two),(Get-Item -LiteralPath $one))
        $exe = Join-Path $TestDrive 'pdftk.exe'
        Mock Write-NativeProcessLog { }
    }

    It 'retains all given inputs and their order, sums pages and rechecks before merge' {
        Mock Get-PdfDocumentInspection {
            param($LiteralPath)
            [pscustomobject]@{ Succeeded=$true; PageCount=if($LiteralPath -eq $two){[long]2}else{[long]1}; InputError=$null; NativeResult=$null }
        }
        $inventory = Get-PdfInputInventory -Executable $exe -Inputs $inputs
        $inventory.Inputs.Count | Should -Be 2
        ($inventory.Inputs.FullName -join '|') | Should -BeExactly ($two + '|' + $one)
        $inventory.ExpectedPageCount | Should -Be 3
        { Assert-PdfInputInventory -Inventory $inventory } | Should -Not -Throw
        [IO.File]::AppendAllText($one,'changed before merge')
        { Assert-PdfInputInventory -Inventory $inventory } | Should -Throw '*changed*'
        Should -Invoke Get-PdfDocumentInspection -Times 2 -Exactly
    }

    It 'stops on a rejected file without returning a partial inventory' {
        Mock Get-PdfDocumentInspection {
            param($LiteralPath)
            [pscustomobject]@{ Succeeded=$false; PageCount=$null; InputError=('controlled named rejection: '+$LiteralPath); NativeResult=$null }
        }
        { Get-PdfInputInventory -Executable $exe -Inputs $inputs } | Should -Throw '*controlled named rejection*'
        Should -Invoke Get-PdfDocumentInspection -Times 1 -Exactly
    }

    It 'detects mutation of the file during native inspection' {
        Mock Get-PdfDocumentInspection {
            param($LiteralPath)
            [IO.File]::AppendAllText($LiteralPath,'changed during inspection')
            [pscustomobject]@{ Succeeded=$true; PageCount=[long]1; InputError=$null; NativeResult=$null }
        }
        { Get-PdfInputInventory -Executable $exe -Inputs $inputs } | Should -Throw '*changed*'
    }

    It 'snapshots every file before the first inspection and detects a later-file mutation' {
        Mock Get-PdfDocumentInspection {
            param($LiteralPath)
            if ($LiteralPath -eq $two) { [IO.File]::AppendAllText($one,'changed while earlier file inspected') }
            [pscustomobject]@{ Succeeded=$true; PageCount=[long]1; InputError=$null; NativeResult=$null }
        }
        { Get-PdfInputInventory -Executable $exe -Inputs $inputs } | Should -Throw '*changed*'
        Should -Invoke Get-PdfDocumentInspection -Times 1 -Exactly
    }

    It 'detects an earlier-file mutation during a later inspection at the final check' {
        Mock Get-PdfDocumentInspection {
            param($LiteralPath)
            if ($LiteralPath -eq $one) { [IO.File]::AppendAllText($two,'changed after its inspection') }
            [pscustomobject]@{ Succeeded=$true; PageCount=[long]1; InputError=$null; NativeResult=$null }
        }
        $inventory = Get-PdfInputInventory -Executable $exe -Inputs $inputs
        { Assert-PdfInputInventory -Inventory $inventory } | Should -Throw '*changed*'
    }

    It 'does not return success if native evidence logging fails' {
        Mock Get-PdfDocumentInspection {
            [pscustomobject]@{ Succeeded=$true; PageCount=[long]1; InputError=$null; NativeResult=(New-InputPreflightResult) }
        }
        Mock Write-NativeProcessLog { throw 'controlled log write failure' }
        { Get-PdfInputInventory -Executable $exe -Inputs $inputs -LogPath (Join-Path $TestDrive 'run.log') } | Should -Throw '*log write failure*'
        Should -Invoke Get-PdfDocumentInspection -Times 1 -Exactly
    }

    It 'returns exactly one inventory object when the native logger emits text' {
        Mock Get-PdfDocumentInspection {
            [pscustomobject]@{ Succeeded=$true; PageCount=[long]1; InputError=$null; NativeResult=(New-InputPreflightResult) }
        }
        Mock Write-NativeProcessLog { 'controlled native diagnostic line' }
        $returned = @(Get-PdfInputInventory -Executable $exe -Inputs $inputs -LogPath (Join-Path $TestDrive 'run.log'))
        $returned.Count | Should -Be 1
        $returned[0].Inputs.Count | Should -Be 2
        $returned[0].ExpectedPageCount | Should -Be 2
        Should -Invoke Write-NativeProcessLog -Times 2 -Exactly
    }

    It 'guards the expected total against Int64 overflow' {
        Mock Get-PdfDocumentInspection {
            [pscustomobject]@{ Succeeded=$true; PageCount=[long]::MaxValue; InputError=$null; NativeResult=$null }
        }
        { Get-PdfInputInventory -Executable $exe -Inputs $inputs } | Should -Throw '*page total*'
    }

    It 'rejects an empty inventory before inspection' {
        Mock Get-PdfDocumentInspection { throw 'Unexpected inspection' }
        { Get-PdfInputInventory -Executable $exe -Inputs @() } | Should -Throw '*input*'
        Should -Invoke Get-PdfDocumentInspection -Times 0 -Exactly
    }
}
