@echo off
rem ===========================================================================
rem  Start Reporter (browser edition)
rem
rem  Double-click this. It finds a Python, makes sure the dependencies are
rem  there, and starts the app. It exists because the documented three-line
rem  setup fails on any machine where Python was installed without "Add to
rem  PATH" ticked, and the four errors PowerShell prints in that case say
rem  nothing useful about the actual problem.
rem
rem  Nothing here installs Python. If none is found it says so plainly.
rem ===========================================================================
setlocal
cd /d "%~dp0"
set "PY="
set "REQS=requirements-web.txt"

rem --- 1. A venv beside this script. The normal case once set up. ------------
if exist ".venv\Scripts\python.exe" (
    set "PY=%~dp0.venv\Scripts\python.exe"
    goto :have_python
)

rem --- 2. An explicit path the operator wrote down --------------------------
rem  For a machine where Python exists but nowhere this script would look: put
rem  the full path to python.exe on the first line of python-path.txt, beside
rem  this file. Explicit beats guessing, and beats hunting through the profile
rem  for some other project's virtual environment.
if exist "python-path.txt" (
    for /f "usebackq delims=" %%P in ("python-path.txt") do (
        if not defined PY if exist "%%~P" set "PY=%%~P"
    )
)
if defined PY goto :have_python

rem --- 3. Whatever is on PATH ------------------------------------------------
rem  Each candidate must actually run: a WindowsApps alias is either a real Store
rem  Python or a stub that only points at the Store, and only running it tells.
for /f "delims=" %%P in ('where python.exe 2^>nul') do (
    if not defined PY "%%P" -c "import sys; sys.exit(sys.version_info < (3, 10))" >nul 2>&1 && set "PY=%%P"
)
if defined PY goto :have_python

for /f "delims=" %%P in ('where py.exe 2^>nul') do (
    if not defined PY set "PY=%%P"
)
if defined PY goto :have_python

rem --- 4. The usual install locations ---------------------------------------
for %%D in (
    "%LOCALAPPDATA%\Programs\Python\Python313"
    "%LOCALAPPDATA%\Programs\Python\Python312"
    "%LOCALAPPDATA%\Programs\Python\Python311"
    "%ProgramFiles%\Python313"
    "%ProgramFiles%\Python312"
    "%ProgramFiles%\Python311"
) do (
    if not defined PY if exist "%%~D\python.exe" set "PY=%%~D\python.exe"
)
if defined PY goto :have_python

echo.
echo   Reporter could not find Python on this machine.
echo.
echo   Install Python 3.12 from https://www.python.org/downloads/ and tick
echo   "Add python.exe to PATH" on the first screen of the installer. Then
echo   double-click this file again.
echo.
echo   Already have Python somewhere unusual? Put the full path to its
echo   python.exe on the first line of a file called python-path.txt, saved
echo   next to this one, and double-click again.
echo.
pause
exit /b 1

:have_python
echo Using Python: %PY%

rem --- Dependencies ----------------------------------------------------------
"%PY%" -c "import pdfplumber, docx, docxtpl, pandas, openpyxl, jinja2" >nul 2>&1
if not errorlevel 1 goto :run

echo.
echo   First run: setting up. This takes a couple of minutes and needs the
echo   internet. It happens once.
echo.
if not exist ".venv\Scripts\python.exe" (
    "%PY%" -m venv .venv
    if errorlevel 1 goto :setup_failed
)
set "PY=%~dp0.venv\Scripts\python.exe"
"%PY%" -m pip install --upgrade pip >nul 2>&1
"%PY%" -m pip install -r "%REQS%"
if errorlevel 1 goto :setup_failed

"%PY%" -c "import pdfplumber, docx, docxtpl, pandas, openpyxl, jinja2" >nul 2>&1
if errorlevel 1 goto :setup_failed

:run
echo.
"%PY%" reporter_web.py %*
set "CODE=%ERRORLEVEL%"
if not "%CODE%"=="0" (
    echo.
    echo   Reporter stopped with an error. The lines above say why.
    pause
)
exit /b %CODE%

:setup_failed
echo.
echo   Setup did not finish. The most common causes:
echo     - no internet access, or a proxy blocking pypi.org
echo     - the Python found above is too old ^(3.11 or newer is needed^)
echo.
echo   Python used: %PY%
echo.
pause
exit /b 1
