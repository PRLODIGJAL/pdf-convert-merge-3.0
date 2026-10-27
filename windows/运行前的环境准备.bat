@echo off
rem ============================================================
rem  PDF convert tool v3.0 - pre-run environment prep (read-only)
rem
rem  WHAT IT DOES
rem    1. is the Python runtime complete
rem    2. which file types this PC can convert, and which program
rem       each family will use
rem    3. CAD - is the CAD found (BY NAME, never by install path),
rem            and which PDF printer does it offer
rem    4. BarTender - is its automation (COM) registered and callable
rem    +  drag ONE file onto this .bat and it also dry-runs that one
rem
rem  Nothing is converted, nothing is written, no original file is
rem  touched.  It only looks.
rem
rem  WHY THERE IS NO PATH IN HERE
rem    "%~dp0" is the folder this .bat lives in, so it always works
rem    no matter where you copy the package to.  Never type a path.
rem
rem  WHY THIS FILE IS PURE ASCII
rem    cmd.exe decodes .bat with the console codepage, which differs
rem    between machines.  All Chinese text is printed by Python
rem    (setup_env.py / seq_core.py), never from here.
rem
rem  No parenthesised blocks on purpose - "if errorlevel" inside a
rem  block is a classic source of surprises.
rem ============================================================
setlocal
pushd "%~dp0"

set "PYEXE="

rem ---- 1) prefer the official "py" launcher from python.org ----
py -3 -c "import sys" >nul 2>&1
if not errorlevel 1 set "PYEXE=py -3"
if defined PYEXE goto ENSURE

rem ---- 2) fall back to "python" on PATH ----
python -c "import sys" >nul 2>&1
if not errorlevel 1 set "PYEXE=python"
if defined PYEXE goto ENSURE

echo.
echo   [!] Python 3 was not found on this computer.
echo.
echo       Install it from:
echo         https://www.python.org/downloads/windows/
echo.
echo       During setup, TICK this box:
echo         [x] Add python.exe to PATH
echo.
echo       Then double-click this file again.
echo.
pause
popd
exit /b 1

:ENSURE
set "PYTHONIOENCODING=utf-8"
set "PYTHONUTF8=1"

rem Always let setup_env.py decide.  It probes the existing runtime and
rem rebuilds it when it is missing OR broken (a .venv copied from another
rem computer looks fine at a glance but its interpreter is a dead link) -
rem so a one-line "if exist" shortcut here would silently let a broken
rem runtime through.  When everything is fine this only takes a second.
echo.
echo   Checking the Python runtime.  If this is the first run it sets it
echo   up now - that may take a minute, please do not close the window.
echo.
%PYEXE% "%~dp0setup_env.py" --setup
if not exist "%~dp0.venv\Scripts\python.exe" goto NOENV

set "VPY=%~dp0.venv\Scripts\python.exe"

echo.
echo ============================================================
echo   [1/3]  Python environment
echo ============================================================
"%VPY%" "%~dp0setup_env.py" --check

echo.
echo ============================================================
echo   [2/3]  Which file types can this PC convert
echo ============================================================
"%VPY%" "%~dp0seq_core.py" --show-rules --brief

echo.
echo ============================================================
echo   [3/3]  CAD
echo   This one starts the CAD once, so it can take 10-40 seconds.
echo ============================================================
"%VPY%" "%~dp0seq_core.py" --show-cad --brief

echo.
echo ============================================================
echo   [4/4]  BarTender
echo   This one starts BarTender once, so it can take a few seconds.
echo ============================================================
"%VPY%" "%~dp0seq_core.py" --show-bartender --brief

if "%~1"=="" goto TAIL

echo.
echo ============================================================
echo   [extra]  Dry run of one file: %~nx1
echo ============================================================
"%VPY%" "%~dp0seq_core.py" --check-convert "%~f1"

:TAIL
echo.
echo ============================================================
echo   Finished. Nothing was converted, no file was changed.
echo   Please copy everything above and send it back.
echo ============================================================
echo.
pause
popd
exit /b 0

:NOENV
echo.
echo   [!] The Python runtime could not be set up.
echo       Send the messages above back so they can be looked at.
echo.
pause
popd
exit /b 1
