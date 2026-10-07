@echo off
setlocal EnableExtensions DisableDelayedExpansion
set "EC=1"

REM Exactly one source folder. Check the raw empty second argument only after
REM its unquoted value has been shown empty, avoiding unquoted metacharacters.
if "%~1"=="" goto usage
if not "%~2"=="" goto usage
if not "%2"=="" goto usage
set "SOURCE=%~1"
if not exist "%SOURCE%\" goto invalid_source

set "PS1=%~dp0WinPDFMerge.ps1"
if not exist "%PS1%" goto missing_script

REM Preserve trailing separators through Windows PowerShell's quoted argv.
REM Only the final backslash run is doubled; the displayed source stays literal.
set "SOURCE_ARG=%SOURCE%"
set "QUOTE_SUFFIX="
:quote_source
if not "%SOURCE_ARG:~-1%"=="\" goto run_script
set "SOURCE_ARG=%SOURCE_ARG:~0,-1%"
set "QUOTE_SUFFIX=%QUOTE_SUFFIX%\\"
goto quote_source

:run_script
REM Remove a local environment shadow of cmd's dynamic ERRORLEVEL value.
set "ERRORLEVEL="
REM Existing process-only policy flag; organization Group Policy still wins.
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%PS1%" -SourceFolder "%SOURCE_ARG%%QUOTE_SUFFIX%"
set "EC=%ERRORLEVEL%"
echo.
if "%EC%"=="0" goto success
if "%EC%"=="2" goto partial_success
echo Merge failed with exit code %EC%.
goto finish

:success
echo Merge completed successfully.
goto finish

:partial_success
echo Partial success ^(exit code 2^). The merged master is retained; check the log.
goto finish

:usage
echo Usage: drag-and-drop exactly one folder containing PDFs onto this file.
goto finish

:invalid_source
echo Provided path is not a folder: "%SOURCE%"
goto finish

:missing_script
echo Can't find WinPDFMerge.ps1 next to this .bat: "%PS1%"

:finish
pause
exit /b %EC%
