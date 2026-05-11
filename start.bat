@echo off
setlocal EnableExtensions
chcp 65001 >nul
cd /d "%~dp0"

rem Optional: set a fixed Python 3.14 path if needed.
rem set "PYTHON_EXE=E:\setups\vmr_sdks\versions\python_versions\python-3.14.0\python.exe"

set "BASE_PY_EXE=E:\setups\vmr_sdks\versions\python_versions\python-3.14.0\python.exe"
rem Backward compatible: if user set BASE_PY_EXE manually, treat it as PYTHON_EXE.
if not defined PYTHON_EXE (
  if defined BASE_PY_EXE (
    set "PYTHON_EXE=%BASE_PY_EXE%"
  )
)

set "BASE_PY_EXE="
call :ResolveBasePython314
if errorlevel 1 (
  echo Could not find Python 3.14+ for this script.
  echo In PowerShell you may see another python than CMD uses.
  echo Please set PYTHON_EXE at the top of this script to your Python 3.14 path.
  pause
  exit /b 1
)

echo Using base Python: %BASE_PY_EXE%
"%BASE_PY_EXE%" --version

echo [1/6] Check Python virtual environment...
if not exist ".venv\Scripts\python.exe" (
  echo .venv not found, creating...
  call :CreateVenv ".venv"
  if errorlevel 1 (
    echo Failed to create virtual environment.
    echo Please set PYTHON_EXE to an absolute Python 3.14 path.
    pause
    exit /b 1
  )
)

set "PYTHON=%CD%\.venv\Scripts\python.exe"

echo Validate .venv Python version...
call :CheckPython314 "%PYTHON%"
if errorlevel 1 (
  echo .venv Python is not 3.14+, rebuilding...
  call :RebuildVenv ".venv"
  if errorlevel 1 (
    echo Failed to rebuild virtual environment.
    echo Base Python path: %BASE_PY_EXE%
    echo Please set PYTHON_EXE to an absolute Python 3.14 path and retry.
    pause
    exit /b 1
  )
  set "PYTHON=%CD%\.venv\Scripts\python.exe"
  call :CheckPython314 "%PYTHON%"
  if errorlevel 1 (
    echo Still not getting Python 3.14+ in .venv.
    echo Please set PYTHON_EXE to an absolute Python 3.14 path.
    pause
    exit /b 1
  )
)

echo [2/6] Auto detect API port...
set "API_PORT="
for /f "usebackq delims=" %%P in (`powershell -NoProfile -ExecutionPolicy Bypass -Command "$ports=8003,8004,8005,8010,8080,9000; $chosen=0; foreach($p in $ports){ try{ $l=[System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback,$p); $l.Start(); $l.Stop(); $chosen=$p; break } catch {} }; Write-Output $chosen"`) do set "API_PORT=%%P"

if "%API_PORT%"=="0" (
  echo No available API port found.
  pause
  exit /b 1
)

if not defined API_PORT (
  echo Failed to detect API port.
  pause
  exit /b 1
)

echo Detected API port: %API_PORT%

echo [3/6] Prepare pip...
"%PYTHON%" -m ensurepip --upgrade >nul 2>&1

echo [4/6] Install/update dependencies...
"%PYTHON%" -m pip install --upgrade pip setuptools wheel
if errorlevel 1 (
  echo Failed to install pip/setuptools/wheel.
  pause
  exit /b 1
)

"%PYTHON%" -m pip install -e .
if errorlevel 1 (
  echo Failed to install project dependencies.
  pause
  exit /b 1
)

echo [5/6] Check .env mail config...
if not exist ".env" (
  echo .env not found, creating now.
  set /p IMAP_USERNAME=Please input IMAP_USERNAME: 
  set /p IMAP_PASSWORD=Please input IMAP_PASSWORD: 
  (
    echo IMAP_USERNAME=%IMAP_USERNAME%
    echo IMAP_PASSWORD=%IMAP_PASSWORD%
  ) > ".env"
  echo .env created.
) else (
  echo .env exists, skip create.
)

echo [6/6] Start API and scheduler...
start "AI Price API" cmd /k ""%PYTHON%" "api_server.py""
start "AI Price Scheduler" cmd /k ""%PYTHON%" "scripts\run_scheduler.py" --interval 300"

echo Started successfully.
echo API: http://127.0.0.1:%API_PORT%/docs
pause
exit /b 0

:CreateVenv
"%BASE_PY_EXE%" -m venv "%~1"
if errorlevel 1 exit /b 1
if not exist "%~1\Scripts\python.exe" exit /b 1
exit /b 0

:RebuildVenv
"%BASE_PY_EXE%" -m venv --clear "%~1"
if errorlevel 1 (
  "%BASE_PY_EXE%" -m venv "%~1"
  if errorlevel 1 exit /b 1
)
if not exist "%~1\Scripts\python.exe" exit /b 1
exit /b 0

:ResolveBasePython314
if defined PYTHON_EXE (
  if exist "%PYTHON_EXE%" (
    call :CheckPython314 "%PYTHON_EXE%"
    if not errorlevel 1 (
      set "BASE_PY_EXE=%PYTHON_EXE%"
      exit /b 0
    )
  )
)

call :ResolveFromPyLauncher
if not errorlevel 1 exit /b 0

for /f "delims=" %%P in ('where python 2^>nul') do (
  call :CheckPython314 "%%P"
  if not errorlevel 1 (
    set "BASE_PY_EXE=%%P"
    exit /b 0
  )
)

exit /b 1

:ResolveFromPyLauncher
where py >nul 2>nul
if errorlevel 1 exit /b 1

set "_PY_PATH_OUT=%TEMP%\py_path_%RANDOM%_%RANDOM%.txt"
py -3.14 -c "import sys; print(sys.executable)" > "%_PY_PATH_OUT%" 2>nul
if errorlevel 1 (
  del /q "%_PY_PATH_OUT%" >nul 2>nul
  exit /b 1
)

set "_PY_CAND="
set /p _PY_CAND=<"%_PY_PATH_OUT%"
del /q "%_PY_PATH_OUT%" >nul 2>nul

if not defined _PY_CAND exit /b 1
if not exist "%_PY_CAND%" exit /b 1

call :CheckPython314 "%_PY_CAND%"
if errorlevel 1 exit /b 1

set "BASE_PY_EXE=%_PY_CAND%"
exit /b 0

:CheckPython314
set "_CHECK_OUT=%TEMP%\pyver_check_%RANDOM%_%RANDOM%.txt"
"%~1" -c "import sys; print(1 if sys.version_info >= (3,14) else 0)" > "%_CHECK_OUT%" 2>nul
if errorlevel 1 (
  del /q "%_CHECK_OUT%" >nul 2>nul
  exit /b 1
)
set "_PY_VER_OK="
set /p _PY_VER_OK=<"%_CHECK_OUT%"
del /q "%_CHECK_OUT%" >nul 2>nul
if "%_PY_VER_OK%"=="1" (
  exit /b 0
)
exit /b 1
