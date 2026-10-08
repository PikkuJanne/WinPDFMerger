# Controlled snapshot regression tests only; no PDF engine or native claim.
BeforeAll {
    $repo = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).ProviderPath
    . (Join-Path $repo 'tests/CorpusSafetySupport.ps1')

    function ConvertFrom-CorpusSnapshot([string]$Snapshot) {
        $parsed = ConvertFrom-Json -InputObject $Snapshot
        foreach ($row in $parsed) { $row }
    }

    function New-CorpusSnapshotCase {
        $root = Join-Path $TestDrive ('corpus-snapshot-' + [Guid]::NewGuid().ToString('N'))
        $nested = Join-Path $root 'nested'
        $empty = Join-Path $root 'empty.pdf'
        foreach ($directory in @($root,$nested,$empty)) { [void][IO.Directory]::CreateDirectory($directory) }
        $inputPath = Join-Path $root 'input.PDF'
        $hidden = Join-Path $root 'hidden.PDF'
        $nestedFile = Join-Path $nested 'nested.PDF'
        foreach ($file in @($inputPath,$hidden,$nestedFile)) { [IO.File]::WriteAllText($file,'AAAA',[Text.Encoding]::ASCII) }
        [IO.File]::SetAttributes($hidden,([IO.File]::GetAttributes($hidden) -bor [IO.FileAttributes]::Hidden))
        [pscustomobject]@{Root=$root; Nested=$nested; Empty=$empty; Input=$inputPath; Hidden=$hidden; NestedFile=$nestedFile}
    }
}

Describe 'Controlled full-tree corpus safety snapshot regressions' {
    It 'includes hidden and nested files and empty directories with file hashes and exact metadata' {
        $case = New-CorpusSnapshotCase
        $rows = @(ConvertFrom-CorpusSnapshot (Get-CorpusSafetyTreeSnapshot @($case.Root)))
        $expectedPaths = @($case.Root,$case.Nested,$case.Empty,$case.Input,$case.Hidden,$case.NestedFile)
        ($rows.Path | Sort-Object) -join "`n" | Should -BeExactly (($expectedPaths | Sort-Object) -join "`n")
        @($rows | Where-Object Kind -eq 'file').Count | Should -Be 3
        @($rows | Where-Object Kind -eq 'directory').Count | Should -Be 3
        foreach ($row in $rows) {
            $item = Get-Item -LiteralPath $row.Path -Force
            $row.Attributes | Should -Be ([int]$item.Attributes)
            $row.CreatedUtcTicks | Should -Be $item.CreationTimeUtc.Ticks
            if ($row.Kind -eq 'file') {
                $row.SHA256 | Should -Match '^[0-9a-f]{64}$'
                $row.Length | Should -Be 4
                $row.ModifiedUtcTicks | Should -Be $item.LastWriteTimeUtc.Ticks
            } else {
                $row.SHA256 | Should -BeNullOrEmpty
                $row.ModifiedUtcTicks | Should -BeNullOrEmpty
            }
        }
    }

    It 'excludes directory last-write settling while retaining directory presence, attributes and creation time' {
        $case = New-CorpusSnapshotCase
        $before = Get-CorpusSafetyTreeSnapshot @($case.Root)
        $directoryBefore = Get-Item -LiteralPath $case.Empty -Force
        [IO.Directory]::SetLastWriteTimeUtc($case.Empty,$directoryBefore.LastWriteTimeUtc.AddMinutes(2))
        (Get-CorpusSafetyTreeSnapshot @($case.Root)) | Should -BeExactly $before
        [IO.File]::SetAttributes($case.Empty,($directoryBefore.Attributes -bor [IO.FileAttributes]::Hidden))
        (Get-CorpusSafetyTreeSnapshot @($case.Root)) | Should -Not -BeExactly $before
        $row = @(ConvertFrom-CorpusSnapshot (Get-CorpusSafetyTreeSnapshot @($case.Root)) | Where-Object Path -eq $case.Empty)[0]
        $row.Kind | Should -BeExactly 'directory'
        $row.CreatedUtcTicks | Should -Be $directoryBefore.CreationTimeUtc.Ticks
    }

    It 'detects a same-length source byte replacement even if the original file modification time is restored' {
        $case = New-CorpusSnapshotCase
        $before = Get-CorpusSafetyTreeSnapshot @($case.Root)
        $modified = [IO.File]::GetLastWriteTimeUtc($case.Input)
        [IO.File]::WriteAllText($case.Input,'BBBB',[Text.Encoding]::ASCII)
        [IO.File]::SetLastWriteTimeUtc($case.Input,$modified)
        $after = Get-CorpusSafetyTreeSnapshot @($case.Root)
        $after | Should -Not -BeExactly $before
        $old = @(ConvertFrom-CorpusSnapshot $before | Where-Object Path -eq $case.Input)[0]
        $new = @(ConvertFrom-CorpusSnapshot $after | Where-Object Path -eq $case.Input)[0]
        $new.Length | Should -Be $old.Length
        $new.ModifiedUtcTicks | Should -Be $old.ModifiedUtcTicks
        $new.SHA256 | Should -Not -BeExactly $old.SHA256
    }

    It 'still detects source file timestamp and attribute changes and added empty-directory inventory' {
        $case = New-CorpusSnapshotCase
        $before = Get-CorpusSafetyTreeSnapshot @($case.Root)
        $modified = [IO.File]::GetLastWriteTimeUtc($case.Input)
        $attributes = [IO.File]::GetAttributes($case.Input)
        [IO.File]::SetLastWriteTimeUtc($case.Input,$modified.AddMinutes(2))
        (Get-CorpusSafetyTreeSnapshot @($case.Root)) | Should -Not -BeExactly $before
        [IO.File]::SetLastWriteTimeUtc($case.Input,$modified)
        [IO.File]::SetAttributes($case.Input,($attributes -bxor [IO.FileAttributes]::Archive))
        (Get-CorpusSafetyTreeSnapshot @($case.Root)) | Should -Not -BeExactly $before
        [IO.File]::SetAttributes($case.Input,$attributes)
        (Get-CorpusSafetyTreeSnapshot @($case.Root)) | Should -BeExactly $before
        [void][IO.Directory]::CreateDirectory((Join-Path $case.Root 'added-empty-directory'))
        (Get-CorpusSafetyTreeSnapshot @($case.Root)) | Should -Not -BeExactly $before
    }
}
