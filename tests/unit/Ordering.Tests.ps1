BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
    . (Join-Path $repo 'src/WinPDFMerge.Helpers.ps1')

    function New-OrderingItem([string]$Name, [string]$FullName = ('C:\Synthetic\' + $Name + '.pdf')) {
        [pscustomobject]@{ BaseName = $Name; FullName = $FullName; OriginalToken = [Guid]::NewGuid().ToString('N') }
    }

    function Get-OrderingNames([string[]]$Names) {
        $items = @($Names | ForEach-Object { New-OrderingItem $_ })
        @(Sort-PdfInputs -Inputs $items | ForEach-Object { $_.BaseName }) -join ','
    }
}

Describe 'AC011: ASCII numeric magnitude and leading-zero order' {
    It 'orders 1, 01, 001, 2 and 10 by the documented run rules' {
        Get-OrderingNames @('10', '001', '2', '01', '1') | Should -BeExactly '1,01,001,2,10'
    }

    It 'handles tokens beyond Int32, Int64 and decimal without overflow' {
        $eightyNines = '9' * 80
        $eightyZerosAfterOne = '1' + ('0' * 80)
        $expected = @('2', '2147483647', '2147483648', '9223372036854775807', '9223372036854775808', $eightyNines, $eightyZerosAfterOne)
        $reversed = @($expected)
        [Array]::Reverse($reversed)
        Get-OrderingNames $reversed | Should -BeExactly ($expected -join ',')
    }

    It 'uses ordinal digits to distinguish very long values of the same significant length' {
        $first = '8' + ('0' * 78) + '1'
        $second = '8' + ('0' * 78) + '2'
        Get-OrderingNames @($second, $first) | Should -BeExactly ($first + ',' + $second)
    }

    It 'orders equal long magnitudes by original numeric-run length' {
        $number = '1' + ('0' * 80)
        Get-OrderingNames @(('000' + $number), ('0' + $number), $number) | Should -BeExactly (@($number, ('0' + $number), ('000' + $number)) -join ',')
    }

    It 'continues through multiple numeric groups' {
        Get-OrderingNames @('chapter10part1', 'chapter1part10', 'chapter2part1', 'chapter1part2') | Should -BeExactly 'chapter1part2,chapter1part10,chapter2part1,chapter10part1'
    }
}

Describe 'AC012: deterministic natural segments and final ties' {
    It 'orders all-zero numeric runs by their original length' {
        Get-OrderingNames @('000', '0', '0000', '00') | Should -BeExactly '0,00,000,0000'
    }

    It 'applies an equal-magnitude run-length tie before comparing later numeric groups' {
        Get-OrderingNames @('a00b1', 'a0b9', 'a000b0') | Should -BeExactly 'a0b9,a00b1,a000b0'
    }

    It 'compares later natural segments before the original case-sensitive name tie' {
        Get-OrderingNames @('A10', 'a2', 'a10', 'A2') | Should -BeExactly 'A2,a2,A10,a10'
    }

    It 'orders exhausted token sequences before extensions despite original-case differences' {
        Get-OrderingNames @('A1b2', 'a1', 'A1b', 'a') | Should -BeExactly 'a,a1,A1b,A1b2'
    }

    It 'compares mixed digit and text punctuation tokens ordinally rather than putting all numbers first' {
        # These are comparator token fixtures, not a filename-support claim.
        $names = @('a', '!', ':', '10', '2', '-')
        Get-OrderingNames $names | Should -BeExactly '!,-,2,10,:,a'
        [Array]::Reverse($names)
        Get-OrderingNames $names | Should -BeExactly '!,-,2,10,:,a'
    }

    It 'uses the ordinal original basename before the ordinal full-path tie' {
        $upper = New-OrderingItem 'Item2' 'C:\Synthetic\Z\Item2.pdf'
        $lower = New-OrderingItem 'item2' 'C:\Synthetic\A\item2.pdf'
        Compare-PdfInput -Left $upper -Right $lower | Should -BeLessThan 0
        [object]::ReferenceEquals((@(Sort-PdfInputs -Inputs @($lower, $upper)))[0], $upper) | Should -BeTrue
    }

    It 'compares already canonical absolute full paths ordinally when basenames tie' {
        $upper = New-OrderingItem 'same' 'C:\Synthetic\A\same.pdf'
        $lower = New-OrderingItem 'same' 'C:\Synthetic\a\same.pdf'
        Compare-PdfInput -Left $upper -Right $lower | Should -BeLessThan 0
        $items = @(Sort-PdfInputs -Inputs @($lower, $upper))
        $items.Count | Should -Be 2
        [object]::ReferenceEquals($items[0], $upper) | Should -BeTrue
        [object]::ReferenceEquals($items[1], $lower) | Should -BeTrue
    }

    It 'returns zero when natural name, original name and canonical full path are identical' {
        $left = New-OrderingItem 'same2'
        $right = New-OrderingItem 'same2'
        Compare-PdfInput -Left $left -Right $right | Should -Be 0
    }

    It 'treats Arabic-Indic and fullwidth digits as ordinal text rather than numeric runs' {
        $arabicTwelve = 'part' + [char]0x0661 + [char]0x0662
        $arabicTwo = 'part' + [char]0x0662
        $fullwidthTwelve = 'part' + [char]0xff11 + [char]0xff12
        $fullwidthTwo = 'part' + [char]0xff12
        $expected = @('part10', $arabicTwelve, $arabicTwo, $fullwidthTwelve, $fullwidthTwo)
        Get-OrderingNames @($fullwidthTwo, $arabicTwo, $fullwidthTwelve, 'part10', $arabicTwelve) | Should -BeExactly ($expected -join ',')
    }

    It 'gives the same documented order and reversed-input result in <Culture>' -TestCases @(
        @{ Culture = 'en-US' },
        @{ Culture = 'tr-TR' }
    ) {
        param($Culture)
        $savedCulture = [Threading.Thread]::CurrentThread.CurrentCulture
        $savedUiCulture = [Threading.Thread]::CurrentThread.CurrentUICulture
        try {
            [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo($Culture)
            [Threading.Thread]::CurrentThread.CurrentUICulture = [Globalization.CultureInfo]::GetCultureInfo($Culture)
            $nonAscii = [string][char]0x00e4 + '2'
            $expected = @('I2', 'i2', 'I10', 'i10', 'z2', $nonAscii)
            $forward = @('i10', $nonAscii, 'I10', 'I2', 'z2', 'i2')
            Get-OrderingNames $forward | Should -BeExactly ($expected -join ',')
            [Array]::Reverse($forward)
            Get-OrderingNames $forward | Should -BeExactly ($expected -join ',')
        } finally {
            [Threading.Thread]::CurrentThread.CurrentCulture = $savedCulture
            [Threading.Thread]::CurrentThread.CurrentUICulture = $savedUiCulture
        }
    }

    It 'preserves object identities and the original collection while sorting a copy' {
        $original = @(New-OrderingItem '10'; New-OrderingItem '2'; New-OrderingItem '1')
        $before = $original | ConvertTo-Json -Depth 4 -Compress
        $sorted = @(Sort-PdfInputs -Inputs $original)
        $sorted.Count | Should -Be 3
        [object]::ReferenceEquals($sorted[0], $original[2]) | Should -BeTrue
        [object]::ReferenceEquals($sorted[1], $original[1]) | Should -BeTrue
        [object]::ReferenceEquals($sorted[2], $original[0]) | Should -BeTrue
        ($original | ConvertTo-Json -Depth 4 -Compress) | Should -BeExactly $before
    }

    It 'normalizes empty and single-input collections without copying their objects' {
        @(Sort-PdfInputs -Inputs @()).Count | Should -Be 0
        $original = New-OrderingItem '1'
        $single = @(Sort-PdfInputs -Inputs @($original))
        $single.Count | Should -Be 1
        [object]::ReferenceEquals($single[0], $original) | Should -BeTrue
    }

    It 'is reflexive, antisymmetric and transitive on a concise mixed corpus' {
        $items = @('!', '-', '0', '00', '1', '01', '2', '10', 'A2', 'a2') | ForEach-Object { New-OrderingItem $_ }
        foreach ($left in $items) {
            Compare-PdfInput -Left $left -Right $left | Should -Be 0
            foreach ($right in $items) {
                $lr = [Math]::Sign([int](Compare-PdfInput -Left $left -Right $right))
                $rl = [Math]::Sign([int](Compare-PdfInput -Left $right -Right $left))
                $lr | Should -Be (-$rl)
                foreach ($last in $items) {
                    if ($lr -le 0 -and (Compare-PdfInput -Left $right -Right $last) -le 0) {
                        Compare-PdfInput -Left $left -Right $last | Should -BeLessOrEqual 0
                    }
                }
            }
        }
    }
}
