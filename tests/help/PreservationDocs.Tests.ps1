# AC045 public documentation contract. These are text/help checks, not PDF fidelity tests.
# Get-Help reads comment help; the application is never invoked or dot-sourced here.
BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
    $entry = Join-Path $repo 'WinPDFMerge.ps1'
    $work = Join-Path $repo ('tests/.work/T19-preservation-docs/' + [Guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($work)
    $observations = New-Object 'System.Collections.Generic.List[object]'
    $surfaces = @{}
    $sourceBindings = New-Object 'System.Collections.Generic.List[object]'
    foreach ($item in @(
        @{Name='README';RelativePath='README.md'},
        @{Name='Limits';RelativePath='docs/PDF_LIMITATIONS.md'},
        @{Name='Entry';RelativePath='WinPDFMerge.ps1'}
    )) {
        $path = Join-Path $repo $item.RelativePath
        $exists = [IO.File]::Exists($path)
        $hash = $null
        $snapshot = $null
        if ($exists) {
            $hash = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
            $surfaces[$item.Name] = [IO.File]::ReadAllText($path)
            $snapshot = Join-Path $work ($item.Name + '-source.txt')
            [IO.File]::Copy($path,$snapshot,$false)
        } else { $surfaces[$item.Name] = '' }
        $sourceBindings.Add([pscustomobject]@{Path=$item.RelativePath;Exists=$exists;SHA256=$hash;RetainedSourcePath=$snapshot})
    }
    $help = Get-Help -Name $entry -Full
    $surfaces.Help = ($help | Out-String -Width 240)
    [IO.File]::WriteAllText((Join-Path $work 'actual-help.txt'),$surfaces.Help,(New-Object Text.UTF8Encoding($false)))
    $helpBinding = [pscustomobject]@{
        Command=@('Get-Help','-Name',$entry,'-Full');RetainedTextPath=(Join-Path $work 'actual-help.txt');
        SHA256=(Get-FileHash -LiteralPath (Join-Path $work 'actual-help.txt') -Algorithm SHA256).Hash.ToLowerInvariant();
        ApplicationInvoked=$false
    }
    function Add-PreservationDocsObservation([string]$Label,[string]$Surface,$Detail) {
        $observations.Add([pscustomobject]@{
            Label=$Label;Surface=$Surface;
            Scope='public documentation/comment-help contract only; no native PDF, structural, visual or signature proof';
            Detail=$Detail
        })
    }
    function Get-PreservationParagraphs([string]$Text) {
        @([regex]::Split($Text,'(?:\r?\n)[ \t]*(?:\r?\n)') | ForEach-Object {
            [regex]::Replace($_,'\s+',' ').Trim()
        } | Where-Object { $_ })
    }
    function Get-UnqualifiedPreservationClaims([string]$Text) {
        $claims = New-Object 'System.Collections.Generic.List[string]'
        foreach ($paragraph in @(Get-PreservationParagraphs $Text)) {
            foreach ($sentence in @([regex]::Split($paragraph,'(?<=[.!?])\s+'))) {
                $claimPattern = '\b(?:lossless|archiv(?:e|al)[- ]safe)\b|\b(?:guarantees?|ensures?|certifies?)\b.{0,90}\b(?:PDF/A|signature|archival|malware|saniti[sz]|forms?|tags?|attachments?|features?)\b|\b(?:all|every|universal(?:ly)?)\b.{0,60}\b(?:form|tag|attachment|feature|signature)s?\b.{0,60}\b(?:preserv|retain|valid|safe|guarantee)'
                $qualificationPattern = '\b(?:not|no|never|neither|cannot|unqualified|unsupported|unvalidated)\b|\b(?:do|does|can|will)\s+not\b|\b(?:without|lack(?:s|ing)?)\s+(?:a\s+)?guarantee'
                if ($sentence -match $claimPattern -and $sentence -notmatch $qualificationPattern) {
                    $claims.Add($sentence)
                }
            }
        }
        $claims.ToArray()
    }
}

Describe 'AC045 conservative public preservation contract' {
    It 'avoids unqualified preservation promises in <Surface>' -TestCases @(
        @{Surface='README'},@{Surface='Help'},@{Surface='Limits'}
    ) {
        param($Surface)
        $text = [string]$surfaces[$Surface]
        $claims = @(Get-UnqualifiedPreservationClaims $text)
        Add-PreservationDocsObservation ('no-overclaim-'+$Surface) $Surface @{UnqualifiedClaims=$claims}
        $text.Length | Should -BeGreaterThan 0
        $claims.Count | Should -Be 0
    }

    It 'distinguishes the master from the rewritten email copy in <Surface>' -TestCases @(
        @{Surface='README'},@{Surface='Help'}
    ) {
        param($Surface)
        $paragraphs = @(Get-PreservationParagraphs ([string]$surfaces[$Surface]))
        $master = @($paragraphs | Where-Object {
            $_ -match '\bmaster\b' -and $_ -match '\bwithout\b.{0,80}\bintentional\b' -and
            $_ -match '\brasteri[sz]ation\b' -and $_ -match '\bdownsampl'
        })
        $email = @($paragraphs | Where-Object {
            $_ -match '\bemail\b' -and $_ -match '\b(?:rewrite|rewrites|rewritten|lossy|loss|appearance|quality|downsampl)'
        })
        Add-PreservationDocsObservation ('master-email-distinction-'+$Surface) $Surface @{MasterParagraphs=$master;EmailParagraphs=$email}
        $master.Count | Should -BeGreaterThan 0
        $email.Count | Should -BeGreaterThan 0
    }

    It 'makes the detailed preservation limits reachable from the README' {
        $links = @([regex]::Matches([string]$surfaces.README,'\[[^\]]+\]\((?:\./)?docs/PDF_LIMITATIONS\.md(?:#[^)]*)?\)'))
        Add-PreservationDocsObservation 'readme-limitations-link' 'README' @{MatchingLinks=@($links | ForEach-Object {$_.Value});LimitsExists=([IO.File]::Exists((Join-Path $repo 'docs/PDF_LIMITATIONS.md')))}
        $links.Count | Should -BeGreaterThan 0
        [IO.File]::Exists((Join-Path $repo 'docs/PDF_LIMITATIONS.md')) | Should -BeTrue
    }

    It 'explains that preserved source files remain the originals to retain' {
        $paragraphs = @(Get-PreservationParagraphs ([string]$surfaces.Limits))
        $source = @($paragraphs | Where-Object {
            $_ -match '\b(?:source|original)s?\b' -and
            $_ -match '\b(?:unchanged|unmodified|untouched|read[- ]only|not\s+(?:modif|edit|overwrit|writ)|never\s+(?:modif|edit|overwrit|writ)|does\s+not\s+(?:modif|edit|overwrit|writ))'
        })
        $retain = @($paragraphs | Where-Object {
            $_ -match '\b(?:keep|retain|preserve)\b.{0,100}\boriginals?\b' -and
            $_ -match '\b(?:signed|signature|feature[- ]rich|forms?|attachments?|tags?)\b'
        })
        Add-PreservationDocsObservation 'retain-source-originals' 'Limits' @{SourceParagraphs=$source;RetentionParagraphs=$retain}
        $source.Count | Should -BeGreaterThan 0
        $retain.Count | Should -BeGreaterThan 0
    }

    It 'states an explicit <Label> limitation' -TestCases @(
        @{Label='signature';Topic='\bsignatures?\b';Extra=''},
        @{Label='XFA';Topic='\bXFA\b';Extra=''},
        @{Label='PDF-A';Topic='\bPDF/A\b';Extra=''},
        @{Label='accessibility-tags';Topic='\b(?:tags?|tagged|accessibility)\b';Extra=''},
        @{Label='attachments';Topic='\battachments?\b|\bembedded\s+files?\b';Extra=''},
        @{Label='forms-repeated-fields';Topic='\bforms?\b|\bAcroForm\b';Extra='\b(?:repeated|duplicate)\b.{0,80}\b(?:fields?|names?)\b'},
        @{Label='malware';Topic='\bmalware\b|\bsaniti[sz](?:er|ation)\b';Extra=''}
    ) {
        param($Label,$Topic,$Extra)
        $paragraphs = @(Get-PreservationParagraphs ([string]$surfaces.Limits))
        $limitationParagraphs = @($paragraphs | Where-Object {
            $_ -match $Topic -and
            $_ -match '\b(?:not|no|never|neither|cannot|unsupported|unvalidated|unguaranteed|lost|loss|invalidat|may)\b|\b(?:can|could)\b.{0,60}\b(?:change|lose|remove|drop|invalidat)|\b(?:do|does|will)\s+not\b'
        })
        Add-PreservationDocsObservation ('explicit-limit-'+$Label) 'Limits' @{MatchingParagraphs=$limitationParagraphs;AdditionalTopicPattern=$Extra}
        $limitationParagraphs.Count | Should -BeGreaterThan 0
        if ($Extra) { ($paragraphs -join ' ') | Should -Match $Extra }
    }
}

AfterAll {
    $receipt = Join-Path $work 'documentation-observations.json'
    $record = [ordered]@{
        Task='T19';Scope='AC045 public documentation contract; no application/native/PDF/manual execution';
        ShellVersion=$PSVersionTable.PSVersion.ToString();ShellEdition=$PSVersionTable.PSEdition;
        SourceBindings=@($sourceBindings.ToArray());HelpBinding=$helpBinding;Observations=@($observations.ToArray())
    }
    [IO.File]::WriteAllText($receipt,($record | ConvertTo-Json -Depth 12),(New-Object Text.UTF8Encoding($false)))
    Write-Host 'Preservation documentation observations:'
    Write-Host ($observations | ConvertTo-Json -Depth 10 -Compress)
    Write-Host ('Preservation documentation receipts: '+$receipt)
}
