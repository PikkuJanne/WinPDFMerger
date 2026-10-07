# Controlled Windows process fixture

Run `tools/test/Build-FakeNative.ps1` from an allowed Windows PowerShell session.
It uses an existing .NET Framework `csc.exe` (or an explicitly supplied
`-CompilerPath`) and returns one absolute executable path. It installs nothing.
Each build gets a unique directory under `tests/.work/fake-native/`, containing
`FakeNative.exe`, `compiler.txt`, and a hash/version receipt `build-info.json`.
The source uses .NET Framework APIs and syntax supported by its C# compiler.

| Arguments after executable | Controlled behavior |
| --- | --- |
| `echo <arguments...>` | JSON array containing every argument after `echo`, including empty values; exit 0. |
| `streams [stdout text] [stderr text]` | One UTF-8 line on each stream; omitted values are `fake stdout` and `fake stderr`; exit 0. |
| `fail <exit 1..255> [absolute normalized partial path]` | Optional create-new file containing a non-PDF marker, then one line on each stream and the requested exit. The harness supplies a run-owned file path whose parent exists. Existing files are preserved. |
| `sleep <milliseconds 0..300000> [absolute new PID receipt path]` | Optionally write the exact fixture PID with create-new semantics, then wait with no stream output; exit 0. A timeout test must terminate only its own fixture process. |
| `flood <lines 0..65536> [padding width 1..1024]` | Numbered `stdout:000000:` and `stderr:000000:` lines with `x` padding. Default padding width 128. Requests exceeding 16 MiB per stream fail before flood output; exit 0 for valid requests. |
| `stdin` | Read stdin until EOF, then report the character count on stdout; verifies the runner closes child stdin. |
| `environment <variable name>` | Report `<unset>` or a JSON single-element value array; observes only this fixture's inherited child environment. |
| `hold-pipes <milliseconds 0..300000> <absolute new child-PID receipt path>` | Start one sleeping child that inherits stdout/stderr, create a receipt with its exact PID, emit both streams and exit immediately. The test owns and terminates that child by its exact PID after inspecting bounded capture; the application runner does not claim descendant termination. |

Invalid modes, bounds, or IO operations report an error on stderr and exit 64.
The fake neither invokes PDFtk/Ghostscript nor creates or validates PDFs. Its
results count only as controlled argument/process/fault tests. The fixture must
not appear on application dependency lookup paths or in the release package.
