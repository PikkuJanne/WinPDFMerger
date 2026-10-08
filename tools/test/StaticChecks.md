# Development static checks

`Invoke-StaticChecks.ps1` uses the parser of its actual host and imports only
PSScriptAnalyzer **1.25.0**, pinned in `tests/TestDependencies.psd1`. It downloads
or installs nothing and changes no execution policy. Select Windows PowerShell
5.1 and the approved portable **PowerShell 7.6.6** separately; PATH alone does
not establish the reference host. The outer acceptance capture verifies cache
provenance and selected binary hashes before invocation.

```powershell
powershell.exe -NoProfile -File tools/test/Invoke-StaticChecks.ps1 -AnalyzerModulePath 'C:\explicit\module\PSScriptAnalyzer.psd1'
& 'C:\explicit\PowerShell-7.6.6\pwsh.exe' -NoProfile -File tools/test/Invoke-StaticChecks.ps1 -AnalyzerModulePath 'C:\explicit\module\PSScriptAnalyzer.psd1'
```

The default scope includes all tracked and nonignored maintained PowerShell
files in the entry script, `src/`, `tests/`, `tools/test/`, and root analyzer
settings. Git enumeration excludes generated ignored `.work` trees. Historical
scripts under `docs/codex/evidence/` are immutable execution receipts and are
outside this maintained-code gate. Both application files are explicitly
required. `-SourcePath` accepts existing individual PowerShell files for focused
checker regressions; its receipts say `explicit-selected-files` and cannot be
mistaken for the whole repository gate.

No selected source is executed or dot-sourced. Each source first receives an
actual `Parser.ParseFile` check. A parse error fails the run and records that
file's analyzer as `not_run`; it is never converted into a skipped or empty pass.
Pinned analysis then checks every severity of every selected rule. Any finding
fails the gate. Current maintained code needs no inline suppressions. The runner
also records selected suppressed findings and fails on any such finding, so an
added suppression cannot silently make this gate green.

One unique `tests/.work/static/<GUID>/analysis.json` contains commit/dirty state,
actual shell/edition/bitness/policy, analyzer version and module-file hashes,
settings/pin hashes, rule/target versions, per-file source hashes/findings and
visible parser/analyzer pass/fail/not_run/skip counts. Source bytes and Git state
are checked again after analysis. Reports use CreateNew UTF-8 writes. Exit 0
requires all selected files and source/checkpoint guards to pass; exit 1 reflects
a finding, parser fault, suppression or guard failure. Module loading/selection
errors terminate before a success report and are preparation failures.

## Selected rules and their purpose

`PSScriptAnalyzerSettings.psd1` explicitly selects 41 rules. It has no
`ExcludeRules`, severity filter, wildcard disable, or source suppression.

| Category | Selected checks | Reason |
| --- | --- | --- |
| Invocation and code interpretation | aliases, InvokeExpression, misleading backticks, reserved names/parameters, built-in cmdlet replacement, empty/nonconstant member invocations, cmdlet parameter correctness | Keep executable selection explicit and catch ambiguous or invalid invocation constructs. |
| State and error handling | automatic-variable assignments, empty catches, global aliases/functions/variables, runspace using scope, ShouldContinue without Force | Preserve host state and visible errors. Test fixture locals must follow the same rule. |
| Parameters and pipeline behavior | switch/mandatory defaults, multiple type attributes, empty help messages, parameter-set/kind consistency, pipeline process/single-value checks, ShouldProcess/SupportsShouldProcess consistency | Catch binding and pipeline contracts that can change outcomes. This does not require adding WhatIf to the application interface. |
| Data comparisons and initialization | null comparisons, accidental assignment/redirection, literal hashtable initialization | Catch silent decision and collection errors. |
| Local processing and credentials | plaintext password/secure-string misuse, username+password pairs, PSCredential type, unencrypted authentication, broken hashes, hardcoded computer names, WMI and deprecated manifest fields | Guard against unsafe credential handling and obsolete/nonlocal patterns if introduced later. |
| PowerShell compatibility | BOM for nonASCII files; enabled UseCompatibleSyntax targeting 5.1 and 7.6 | Preserve source decoding in the legacy host and detect newer syntax even when analyzing from PowerShell 7. Actual host parsing remains an independent check. |

The separate vendor-default analysis keeps all ordinary findings visible as
**advisory**, with complete per-file diagnostics and severity counts. It does
not silently label the project free of every vendor rule. Rules about console
Write-Host, plural/helper names, approved verbs, missing OutputType attributes,
formatting or generic state-changing-function ShouldProcess are advisory: the
established local console workflow and internal helpers intentionally use those
patterns. Converting console output to pipeline return objects or expanding the
public WhatIf interface would change behavior outside T22. Pester parameters and
variables shared across BeforeAll/It/AfterAll blocks also cause unused-variable
and unused-parameter diagnostics because ordinary static scope analysis does not
model Pester's block lifecycle. These findings remain disclosed rather than
suppressed in source.

Full command/type compatibility profiles are not selected, and no full
command/type compatibility pass is claimed: the vendor's bundled platform
profiles do not establish the exact required PS7.6.6 runtime or Pester commands,
and static inference cannot prove .NET/native API availability.
Real Windows tests in both required hosts supply that evidence. A static pass
proves neither a successful PDF merge nor native engine or manual acceptance.

`tests/static/StaticChecks.Tests.ps1` exercises no-execution parsing, syntax
failure/not_run counts, shell evaluation, automatic-variable assignment, empty
catches, BOM decoding requirements, advisory visibility, suppressed findings and
mismatched analyzer versions. Existing affected test suites remain the regression
for the test-local renames and receipt retry diagnostic.

Primary behavior references: Microsoft's
[settings, parser diagnostics and suppressions](https://learn.microsoft.com/en-us/powershell/utility-modules/psscriptanalyzer/using-scriptanalyzer?view=ps-modules),
[syntax target configuration](https://learn.microsoft.com/en-us/powershell/utility-modules/psscriptanalyzer/rules/usecompatiblesyntax?view=ps-modules),
and the approved pinned vendor module's `Settings/` and rule catalogue.
