# AC046/AC047 support: public documentation and isolated real-parameter binding.
# This suite never invokes/dot-sources application orchestration or a PDF engine.
BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
    $work = Join-Path $repo ('tests/.work/T20-public-docs/' + [Guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($work)
    $surfaces = @{}
    $bindings = New-Object 'System.Collections.Generic.List[object]'
    $observations = New-Object 'System.Collections.Generic.List[object]'
    foreach ($relative in @('README.md','docs/USAGE.md','docs/TROUBLESHOOTING.md','docs/DEPENDENCIES.md','docs/COMPATIBILITY.md','SECURITY.md','docs/PDF_LIMITATIONS.md','docs/EMAIL_PRESETS.md','LICENSE','WinPDFMerge.ps1','WinPDFMerge.bat')) {
        $path = Join-Path $repo $relative
        $exists = [IO.File]::Exists($path)
        $hash = $null
        if ($exists) {
            $surfaces[$relative] = [IO.File]::ReadAllText($path)
            $hash = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
            [IO.File]::Copy($path,(Join-Path $work ($relative.Replace('/','-') + '.source.txt')),$false)
        } else { $surfaces[$relative] = '' }
        $bindings.Add([pscustomobject]@{Path=$relative;Exists=$exists;SHA256=$hash})
    }
    $tokens = $null; $errors = $null
    $entryAst = [Management.Automation.Language.Parser]::ParseFile((Join-Path $repo 'WinPDFMerge.ps1'),[ref]$tokens,[ref]$errors)
    if ($errors.Count -ne 0 -or $null -eq $entryAst.ParamBlock) { throw 'The real entry parameter block must parse.' }
    # Only the actual param block is extracted. No entry statements are evaluated.
    $bindingAttributes = @($entryAst.ParamBlock.Attributes | ForEach-Object { $_.Extent.Text }) -join "`n"
    $bindingOnly = [scriptblock]::Create($bindingAttributes + "`n" + $entryAst.ParamBlock.Extent.Text + @'

[pscustomobject]@{
    SourceFolder=$SourceFolder;OutputFolder=$OutputFolder;
    SkipEmail=[bool]$SkipEmail;EmailPreset=$EmailPreset
}
'@)
    function Add-PublicDocsObservation([string]$Label,$Detail) {
        $observations.Add([pscustomobject]@{Label=$Label;Detail=$Detail})
    }
    function Get-PublicParagraphs([string]$Text) {
        @([regex]::Split($Text,'(?:\r?\n)[ \t]*(?:\r?\n)') | ForEach-Object { [regex]::Replace($_,'\s+',' ').Trim() } | Where-Object { $_ })
    }
    function Test-PublicParagraph([string]$Text,[string[]]$Patterns) {
        foreach ($paragraph in @(Get-PublicParagraphs $Text)) {
            $matchesAll = $true
            foreach ($pattern in $Patterns) { if ($paragraph -notmatch $pattern) { $matchesAll = $false; break } }
            if ($matchesAll) { return $true }
        }
        return $false
    }
    function Get-PublicLiteral($Element) {
        if ($Element -is [Management.Automation.Language.StringConstantExpressionAst]) { return $Element.Value }
        if ($Element -is [Management.Automation.Language.ConstantExpressionAst]) { return [string]$Element.Value }
        throw ('Public app examples must use literal operands: ' + $Element.Extent.Text)
    }
    function Convert-PublicAppCommand($Command) {
        $name = $Command.GetCommandName()
        if (-not $name) { return $null }
        $elements = @($Command.CommandElements)
        $start = 1
        if ($name -match '(?i)(?:^|[\\/])(?:powershell|pwsh)(?:\.exe)?$') {
            $fileIndex = -1
            for ($index=1; $index -lt $elements.Count; $index++) {
                if ($elements[$index] -is [Management.Automation.Language.CommandParameterAst] -and $elements[$index].ParameterName -eq 'File') { $fileIndex=$index; break }
            }
            if ($fileIndex -lt 0 -or $fileIndex + 1 -ge $elements.Count) { return $null }
            $scriptName = Get-PublicLiteral $elements[$fileIndex+1]
            if ($scriptName -notmatch '(?i)(?:^|[\\/])WinPDFMerge\.ps1$') { return $null }
            $start = $fileIndex + 2
        } elseif ($name -notmatch '(?i)(?:^|[\\/])WinPDFMerge\.ps1$') { return $null }
        $named = @{}
        $positional = New-Object 'System.Collections.Generic.List[string]'
        for ($index=$start; $index -lt $elements.Count; $index++) {
            $element = $elements[$index]
            if ($element -is [Management.Automation.Language.CommandParameterAst]) {
                $parameter = $element.ParameterName
                if ($parameter -eq 'SkipEmail') { $named[$parameter] = $true; continue }
                if ($null -ne $element.Argument) { $named[$parameter] = Get-PublicLiteral $element.Argument; continue }
                if ($index+1 -ge $elements.Count -or $elements[$index+1] -is [Management.Automation.Language.CommandParameterAst]) { throw ('Missing public example operand: ' + $parameter) }
                $index++; $named[$parameter] = Get-PublicLiteral $elements[$index]
            } else { $positional.Add((Get-PublicLiteral $element)) }
        }
        [pscustomobject]@{Text=$Command.Extent.Text;Named=$named;Positional=$positional.ToArray()}
    }
    $examples = New-Object 'System.Collections.Generic.List[object]'
    foreach ($relative in @('README.md','docs/USAGE.md')) {
        foreach ($fence in [regex]::Matches([string]$surfaces[$relative],'(?ms)^```powershell[^\r\n]*\r?\n(.*?)^```')) {
            $exampleTokens=$null; $exampleErrors=$null
            $ast = [Management.Automation.Language.Parser]::ParseInput($fence.Groups[1].Value,[ref]$exampleTokens,[ref]$exampleErrors)
            if ($exampleErrors.Count) { throw ('Public PowerShell fence must parse: '+$relative) }
            foreach ($command in @($ast.FindAll({param($node) $node -is [Management.Automation.Language.CommandAst]},$true))) {
                $converted = Convert-PublicAppCommand $command
                if ($null -ne $converted) { $examples.Add([pscustomobject]@{Path=$relative;Command=$converted}) }
            }
        }
    }
    . (Join-Path $repo 'src/WinPDFMerge.Helpers.ps1')
}

Describe 'AC046 usable public instructions' {
    It 'keeps the complete installation layout and established launchers visible' {
        $readme = [string]$surfaces['README.md']
        foreach ($required in @('WinPDFMerge\.ps1','WinPDFMerge\.bat','src[/\\]WinPDFMerge\.Helpers\.ps1')) { $readme | Should -Match $required }
        $readme | Should -Match '\bWinPDFMerger\b'
        Add-PublicDocsObservation 'complete-layout' @{Launchers=@('WinPDFMerge.ps1','WinPDFMerge.bat');Helper='src/WinPDFMerge.Helpers.ps1'}
    }
    It 'makes optional Ghostscript and explicit master-only use clear' {
        $readme = [string]$surfaces['README.md']
        (Test-PublicParagraph $readme @('\bGhostscript\b','\boptional\b')) | Should -BeTrue
        $readme | Should -Match '-SkipEmail'
    }
    It 'states top-level visible input and retained defaults without a blanket platform claim' {
        $public = [string]$surfaces['README.md'] + "`n`n" + [string]$surfaces['docs/USAGE.md']
        $public | Should -Match '\btop[- ]level\b'
        $public | Should -Match '\bhidden\b'
        (Test-PublicParagraph $public @('\bsubfolders\b','\b(?:not|never|excluded)\b')) | Should -BeTrue
        (Test-PublicParagraph $public @('\bdefault\b','\bscreen\b')) | Should -BeTrue
        (Test-PublicParagraph $public @('\bdefault\b','\b(?:entry[- ]script|scripts?|launcher)\b','\b(?:next\s+to|beside|directory|folder)\b')) | Should -BeTrue
        (Test-PublicParagraph $public @('\bWindows\s+10\b','\b(?:unvalidated|not\s+validated|untested)\b')) | Should -BeTrue
    }
    It 'keeps host OS support separate from scoped dependency test observations' {
        $dependencies = [string]$surfaces['docs/DEPENDENCIES.md']
        (Test-PublicParagraph $dependencies @('Windows.*support\s+channel','\bunestablished\b','Windows\s+PowerShell','\blifecycle\b')) | Should -BeTrue
        (Test-PublicParagraph $dependencies @('\bobservations\b','\bdo\s+not\s+certify\b','\bOS\s+support\b')) | Should -BeTrue
    }
    It 'publishes reasoned Windows 10, live UNC and other-host validation exclusions' {
        $compatibility = [string]$surfaces['docs/COMPATIBILITY.md']
        $compatibility | Should -Match 'Scope\s+recorded\s+\d{4}-\d{2}-\d{2}'
        foreach ($scope in @('Windows\s+10','Live\s+UNC','Windows\s+on\s+ARM','32-bit\s+hosts')) {
            $rows = @($compatibility -split '\r?\n' | Where-Object { $_ -match ('^\|\s*' + $scope + '\s*\|') })
            $rows.Count | Should -Be 1
            $rows[0] | Should -Match 'Excluded\s+from\s+validated\s+v1\.0\.0\s+support'
            $rows[0] | Should -Match 'No\s+actual\s+.+\b(?:test|testing|evidence)'
        }
        (Test-PublicParagraph $compatibility @('\bUNC\b','\bunit\b','\bdo\s+not\s+establish\b','\blive\b')) | Should -BeTrue
        (Test-PublicParagraph $compatibility @('\bx86\b','\bx64\b','\bdo\s+not\s+validate\b','\b32-bit\s+host\b')) | Should -BeTrue
        (Test-PublicParagraph $compatibility @('\bexclusions?\b','\b(?:are|is)\s+not\b','\bpass(?:es|ing)?\b')) | Should -BeTrue
        ([string]$surfaces['README.md']) | Should -Match '\[compatibility\s+scope\]\(docs/COMPATIBILITY\.md\)'
        Add-PublicDocsObservation 'optional-compatibility-exclusions' @{Scopes=@('Windows 10','Live UNC','Windows on ARM','32-bit hosts');ApplicationInvoked=$false;NativeInvoked=$false;ManualAcceptance=$false}
    }
    It 'retains dated official dependency notices and an explicitly acceptable unsigned release' {
        $dependencies = [string]$surfaces['docs/DEPENDENCIES.md']
        $dependencies | Should -Match 'Vendor\s+information\s+checked\s+\d{4}-\d{2}-\d{2}'
        foreach ($official in @('https://github\.com/PowerShell/Announcements/','https://www\.ghostscript\.com/releases/cve/','https://www\.pdflabs\.com/docs/pdftk-license/')) { $dependencies | Should -Match $official }
        (Test-PublicParagraph ([string]$surfaces['SECURITY.md']) @('\bunsigned\b','\bacceptable\b','\boptional\b')) | Should -BeTrue
    }
    It 'parses and binds every public app example to the actual isolated parameter block' {
        $examples.Count | Should -BeGreaterThan 5
        $results = New-Object 'System.Collections.Generic.List[object]'
        foreach ($example in $examples) {
            $named=$example.Command.Named; $positionals=$example.Command.Positional
            $bound = & $bindingOnly @named @positionals
            $bound.SourceFolder | Should -Not -BeNullOrEmpty
            $bound.EmailPreset | Should -BeIn @('screen','ebook')
            $results.Add([pscustomobject]@{Path=$example.Path;Text=$example.Command.Text;Bound=$bound})
        }
        Add-PublicDocsObservation 'isolated-real-parameter-binding' @{Count=$examples.Count;Examples=$results.ToArray();ApplicationInvoked=$false;NativeInvoked=$false}
    }
    It 'rejects an unsupported public option in the binding control' {
        { & $bindingOnly -SourceFolder 'C:\T20 synthetic input' -Recursive } | Should -Throw
    }
    It 'rejects an invalid preset in the binding control' {
        { & $bindingOnly -SourceFolder 'C:\T20 synthetic input' -EmailPreset printer } | Should -Throw
    }
    It 'keeps README default, output, ebook and master-only examples distinct' {
        $readmeExamples = @($examples.ToArray() | Where-Object Path -eq 'README.md')
        $readmeExamples.Count | Should -Be 5
        @($readmeExamples | Where-Object { $_.Command.Named.ContainsKey('OutputFolder') }).Count | Should -Be 1
        @($readmeExamples | Where-Object { $_.Command.Named.ContainsKey('EmailPreset') -and $_.Command.Named['EmailPreset'] -eq 'ebook' }).Count | Should -Be 1
        @($readmeExamples | Where-Object { $_.Command.Named.ContainsKey('SkipEmail') }).Count | Should -Be 1
    }
    It 'documents percent-token recovery with a direct PowerShell route' {
        $public = [string]$surfaces['README.md'] + "`n`n" + [string]$surfaces['docs/TROUBLESHOOTING.md']
        (Test-PublicParagraph $public @('%(?:NAME|[^%\s]+)%','\b(?:expand|expansion)\b','\b(?:cmd|batch|launcher)\b')) | Should -BeTrue
        @($examples.ToArray() | Where-Object { $_.Command.Text -match 'powershell\.exe' -and $_.Command.Text -match '%[^%]+%' }).Count | Should -BeGreaterThan 0
    }
    It 'matches documented success code to the real master-only outcome' {
        $outcome = Get-PdfMergeOutcome -MasterPublished $true -EmailState skipped -MasterPath 'C:\T20 synthetic output\master.pdf'
        $outcome.ExitCode | Should -Be 0
        $rows=@(([string]$surfaces['README.md']) -split '\r?\n' | Where-Object { $_ -match '^\|\s*`?0`?\s*\|' })
        $rows.Count | Should -Be 1
        $rows[0] | Should -Match '\b(?:success|successful)\b'
        $rows[0] | Should -Match '\bmaster\b'
        $rows[0] | Should -Match '\b(?:skip|absent|unavailable|missing|benefit)'
    }
    It 'matches documented failure code to the real pre-master outcome' {
        $outcome = Get-PdfMergeOutcome -MasterPublished $false -EmailState not_started -RunFailed
        $outcome.ExitCode | Should -Be 1
        $rows=@(([string]$surfaces['README.md']) -split '\r?\n' | Where-Object { $_ -match '^\|\s*`?1`?\s*\|' })
        $rows.Count | Should -Be 1
        $rows[0] | Should -Match '\b(?:failure|failed|fail)\b'
    }
    It 'matches documented partial success to the real retained-master outcome' {
        $outcome = Get-PdfMergeOutcome -MasterPublished $true -EmailState failed -MasterPath 'C:\T20 synthetic output\master.pdf' -RunFailed
        $outcome.ExitCode | Should -Be 2
        $rows=@(([string]$surfaces['README.md']) -split '\r?\n' | Where-Object { $_ -match '^\|\s*`?2`?\s*\|' })
        $rows.Count | Should -Be 1
        $rows[0] | Should -Match '\bpartial\b'
        $rows[0] | Should -Match '\bmaster\b'
        $rows[0] | Should -Match '\b(?:retain|kept|safe|surviv)'
    }
    It 'keeps public local links inside the repository and resolves their targets' {
        $checked = New-Object 'System.Collections.Generic.List[object]'
        foreach ($relative in @('README.md','docs/USAGE.md','docs/TROUBLESHOOTING.md','docs/DEPENDENCIES.md','docs/COMPATIBILITY.md','SECURITY.md','docs/PDF_LIMITATIONS.md','docs/EMAIL_PRESETS.md')) {
            $text = [string]$surfaces[$relative]
            $text.Length | Should -BeGreaterThan 0
            foreach ($link in [regex]::Matches($text,'\[[^\]]+\]\(([^)\s]+)(?:\s+"[^"]*")?\)')) {
                $target=$link.Groups[1].Value
                if ($target -match '^[a-z][a-z0-9+.-]*:') { continue }
                $parts=$target -split '#',2
                $pathPart=[Uri]::UnescapeDataString($parts[0])
                $path=if($pathPart){[IO.Path]::GetFullPath((Join-Path (Split-Path -Parent (Join-Path $repo $relative)) $pathPart))}else{Join-Path $repo $relative}
                $path.StartsWith($repo+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase) | Should -BeTrue
                [IO.File]::Exists($path) | Should -BeTrue
                if ($parts.Count -eq 2 -and $parts[1]) {
                    $slugs=New-Object 'System.Collections.Generic.List[string]'
                    foreach ($heading in [regex]::Matches([IO.File]::ReadAllText($path),'(?m)^#{1,6}\s+(.+?)\s*#*\s*$')) {
                        $slug=$heading.Groups[1].Value.ToLowerInvariant() -replace '[^\p{L}\p{N}_ -]','' -replace ' ','-'
                        $slugs.Add($slug)
                    }
                    [Uri]::UnescapeDataString($parts[1]) | Should -BeIn $slugs.ToArray()
                }
                $checked.Add([pscustomobject]@{From=$relative;Target=$target})
            }
        }
        $checked.Count | Should -BeGreaterThan 5
        Add-PublicDocsObservation 'public-local-links' @{Count=$checked.Count;Links=$checked.ToArray()}
    }
}

Describe 'AC047 licensing, privacy and scope' {
    It 'preserves the original MIT license bytes' {
        (Get-FileHash -LiteralPath (Join-Path $repo 'LICENSE') -Algorithm SHA256).Hash.ToLowerInvariant() | Should -Be '714ffa7a21614e637d7dbb17a2e86e4575d7ecd2b6d36b5b67ab3fdcc4193477'
    }
    It 'separates project MIT terms from externally installed vendor licenses' {
        $dependencies=[string]$surfaces['docs/DEPENDENCIES.md']
        $dependencies | Should -Match '\bMIT\b'
        $dependencies | Should -Match '\bPDFtk\b'
        $dependencies | Should -Match '\bGhostscript\b'
        (Test-PublicParagraph $dependencies @('\b(?:dependencies|vendors?|third[- ]party|PDFtk|Ghostscript)\b','\b(?:separate|own|not\s+covered|not\s+relicense|does\s+not\s+relicense)\b','\b(?:licenses?|terms|MIT)\b')) | Should -BeTrue
    }
    It 'discloses local sensitive logs and sanitizing a copy before public sharing' {
        $security=[string]$surfaces['SECURITY.md']
        (Test-PublicParagraph $security @('\blogs?\b','\bpaths?\b','\b(?:names?|metadata)\b')) | Should -BeTrue
        (Test-PublicParagraph $security @('\blogs?\b','\b(?:not|neither|no)\b','\b(?:encrypt|redact)')) | Should -BeTrue
        (Test-PublicParagraph $security @('\b(?:sanitize|redact|replace|remove)\b','\bcopy\b','\b(?:share|sharing|public|report)\b')) | Should -BeTrue
        (Test-PublicParagraph $security @('\b(?:do\s+not|never)\b.{0,80}\b(?:upload|attach|send|publish)\b','\b(?:private|personal|confidential)\b.{0,40}\bPDFs?\b')) | Should -BeTrue
    }
    It 'states offline application processing without download, telemetry or trust promises' {
        $security=[string]$surfaces['SECURITY.md']
        (Test-PublicParagraph $security @('\b(?:application|script|processing)\b','\b(?:local|offline)\b')) | Should -BeTrue
        (Test-PublicParagraph $security @('\b(?:no|not|never)\b.{0,60}\btelemetry\b')) | Should -BeTrue
        (Test-PublicParagraph $security @('\b(?:does\s+not|no|never)\b.{0,80}\b(?:download|install|network)')) | Should -BeTrue
    }
    It 'discloses unsigned status and does not instruct persistent policy bypass' {
        $security=[string]$surfaces['SECURITY.md']
        $security | Should -Match '\bunsigned\b'
        $public=(@('README.md','docs/USAGE.md','docs/TROUBLESHOOTING.md','docs/DEPENDENCIES.md','SECURITY.md') | ForEach-Object { [string]$surfaces[$_] }) -join "`n`n"
        $unsafe=@([regex]::Matches($public,'(?im)^\s*(?:PS[^>]*>\s*)?Set-ExecutionPolicy\b[^\r\n]*'))
        $unsafe.Count | Should -Be 0
        (Test-PublicParagraph $public @('\b(?:Group\s+Policy|organization(?:al)?|enterprise)\b','\b(?:do\s+not|never|cannot|does\s+not|respect|follow|not\s+override)\b')) | Should -BeTrue
    }
    It 'makes the local command-line scope explicit' {
        $readme=[string]$surfaces['README.md']
        $readme | Should -Match '\b(?:website|web\s+service)\b'
        $readme | Should -Match '\b(?:cloud|upload)'
        (Test-PublicParagraph $readme @('\b(?:website|web\s+service|cloud)\b','\b(?:no|not|without|out\s+of\s+scope|outside\s+(?:its|the|project)\s+scope)\b')) | Should -BeTrue
    }
}

AfterAll {
    $record=[ordered]@{
        Task='T20';Scope='AC046/AC047 public-doc contract, isolated real ParamBlock binding and controlled helper outcomes; no application/native/manual execution';
        ShellVersion=$PSVersionTable.PSVersion.ToString();ShellEdition=$PSVersionTable.PSEdition;
        SourceBindings=$bindings.ToArray();Observations=$observations.ToArray()
    }
    [IO.File]::WriteAllText((Join-Path $work 'documentation-observations.json'),($record | ConvertTo-Json -Depth 14),(New-Object Text.UTF8Encoding($false)))
    Write-Host ('Public documentation receipts: '+$work)
}
