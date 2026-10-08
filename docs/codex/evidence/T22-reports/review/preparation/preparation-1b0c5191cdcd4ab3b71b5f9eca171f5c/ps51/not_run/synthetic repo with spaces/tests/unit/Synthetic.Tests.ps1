Describe 'T22 controlled setup fault' { BeforeAll { throw 'T22 synthetic setup fault' }; It 'cannot run' { 1 | Should -Be 1 } }
