@echo off
setlocal EnableDelayedExpansion

:: ── DotnetPOI Phase 11 — Windows Office evidence runner ──────────────────────
::
:: Prerequisites:
::   pip install pywin32 Pillow
::   Microsoft Office must be installed and licensed.
::
:: Usage:
::   run-windows-office-evidence.bat
::
:: Output:
::   tools\manual-verification\evidence\v<version>-<revision>-windows\INDEX.md
:: ─────────────────────────────────────────────────────────────────────────────

set "SCRIPT_DIR=%~dp0"
set "REPO_ROOT=%SCRIPT_DIR%..\..\.."

:: Read project version from csproj
set "DOTNETPOI_VERSION="
for /f "tokens=2 delims=<>" %%v in (
    'findstr /i "VersionPrefix" "%REPO_ROOT%\src\DotnetPoi.Core\DotnetPoi.Core.csproj" 2^>nul'
) do (
    if not defined DOTNETPOI_VERSION set "DOTNETPOI_VERSION=%%v"
)
if not defined DOTNETPOI_VERSION set "DOTNETPOI_VERSION=unknown"

:: Read git revision
set "DOTNETPOI_REVISION="
for /f %%r in ('git -C "%REPO_ROOT%" rev-parse --short HEAD 2^>nul') do (
    set "DOTNETPOI_REVISION=%%r"
)
if not defined DOTNETPOI_REVISION set "DOTNETPOI_REVISION=nogit"

set "DOTNETPOI_EVIDENCE_ID=v%DOTNETPOI_VERSION%-%DOTNETPOI_REVISION%-windows"

echo.
echo  DotnetPOI Phase 11 — Windows Office Evidence
echo  -----------------------------------------------
echo  Version  : %DOTNETPOI_VERSION%
echo  Revision : %DOTNETPOI_REVISION%
echo  Output   : tools\manual-verification\evidence\%DOTNETPOI_EVIDENCE_ID%\INDEX.md
echo.

python "%SCRIPT_DIR%run_windows_office_evidence.py"
set "EXIT_CODE=%ERRORLEVEL%"

echo.
if %EXIT_CODE% equ 0 (
    echo  Result: PASS ^(or PASS_WITH_MISSING_FIXTURES^)
) else if %EXIT_CODE% equ 3 (
    echo  Result: PREREQUISITES MISSING — install pywin32 and Pillow
) else (
    echo  Result: FAIL — see session.log for details
)
echo.

endlocal
exit /b %EXIT_CODE%
