@echo off
setlocal ENABLEDELAYEDEXPANSION

REM Drag & drop a FOLDER onto this .bat
if "%~1"=="" (
  echo Usage: drag-and-drop a folder containing PDFs onto this file.
  pause
  exit /b 1
)

REM Validate folder
if not exist "%~1\" (
  echo Provided path is not a folder: %~1
  pause
  exit /b 1
)

REM Run the PowerShell script from the same directory as this .bat
set SCRIPT_DIR=%~dp0
set PS1=%SCRIPT_DIR%WinPDFMerge.ps1

if not exist "%PS1%" (
  echo Can't find WinPDFMerge.ps1 next to this .bat: %PS1%
  pause
  exit /b 1
)

REM Allow execution, run, and keep window open for status
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS1%" -SourceFolder "%~1"
set EC=%ERRORLEVEL%

echo.
if %EC% EQU 0 (
  echo Merge completed successfully.
) else (
  echo Merge failed with exit code %EC%.
)
pause
exit /b %EC%
