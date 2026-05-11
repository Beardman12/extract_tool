@echo off
setlocal EnableExtensions
chcp 65001 >nul
cd /d "%~dp0"

rem 可选：如果 python 不在 PATH，把路径写在这里
rem set "PYTHON_EXE=C:\Users\YourName\AppData\Local\Programs\Python\Python314\python.exe"

set "BASE_PY_MODE="
set "BASE_PY_EXE="

if defined PYTHON_EXE (
  if exist "%PYTHON_EXE%" (
    set "BASE_PY_MODE=custom"
    set "BASE_PY_EXE=%PYTHON_EXE%"
  )
)

if not defined BASE_PY_MODE (
  where python >nul 2>nul
  if not errorlevel 1 set "BASE_PY_MODE=python"
)

if not defined BASE_PY_MODE (
  where py >nul 2>nul
  if not errorlevel 1 set "BASE_PY_MODE=py"
)

if not defined BASE_PY_MODE (
  echo 未找到可用 Python。
  echo 方案1：安装 Python 3.14+ 并勾选 Add python.exe to PATH
  echo 方案2：在本脚本顶部配置 PYTHON_EXE 为 python.exe 的绝对路径
  pause
  exit /b 1
)

echo [1/6] 检查 Python 虚拟环境...
if not exist ".venv\Scripts\python.exe" (
  echo 未检测到 .venv，开始创建...
  call :RunBasePython -m venv .venv
  if errorlevel 1 (
    echo 创建虚拟环境失败。
    pause
    exit /b 1
  )
)

set "PYTHON=%CD%\.venv\Scripts\python.exe"

echo [2/6] 自动检测 API 端口...
set "API_PORT="
for /f "usebackq delims=" %%P in (`powershell -NoProfile -ExecutionPolicy Bypass -Command "$ports=8003,8004,8005,8010,8080,9000; $chosen=0; foreach($p in $ports){ try{ $l=[System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback,$p); $l.Start(); $l.Stop(); $chosen=$p; break } catch {} }; Write-Output $chosen"`) do set "API_PORT=%%P"

if "%API_PORT%"=="0" (
  echo 未找到可用 API 端口，请关闭占用端口的进程后重试。
  pause
  exit /b 1
)

if not defined API_PORT (
  echo 端口探测失败，请检查 Python 执行环境。
  pause
  exit /b 1
)

echo 检测到可用 API 端口: %API_PORT%

echo [3/6] 准备 pip...
"%PYTHON%" -m ensurepip --upgrade >nul 2>&1

echo [4/6] 安装/更新依赖...
"%PYTHON%" -m pip install --upgrade pip setuptools wheel
if errorlevel 1 (
  echo pip 基础工具安装失败。
  pause
  exit /b 1
)

"%PYTHON%" -m pip install -e .
if errorlevel 1 (
  echo 项目依赖安装失败。
  pause
  exit /b 1
)

echo [5/6] 检查 .env 邮箱配置...
if not exist ".env" (
  echo 未找到 .env，开始引导创建。
  set /p IMAP_USERNAME=请输入 IMAP_USERNAME:
  set /p IMAP_PASSWORD=请输入 IMAP_PASSWORD:
  (
    echo IMAP_USERNAME=%IMAP_USERNAME%
    echo IMAP_PASSWORD=%IMAP_PASSWORD%
  ) > ".env"
  echo .env 已创建。
) else (
  echo 已存在 .env，跳过创建。
)

echo [6/6] 启动 API 和定时服务...
start "AI Price API" cmd /k "set API_PORT=%API_PORT% && \"%PYTHON%\" api_server.py"
start "AI Price Scheduler" cmd /k ""%PYTHON%" scripts\run_scheduler.py --interval 300"

echo 启动完成。
echo API: http://127.0.0.1:%API_PORT%/docs
pause
exit /b 0

:RunBasePython
if "%BASE_PY_MODE%"=="custom" (
  "%BASE_PY_EXE%" %*
  exit /b %errorlevel%
)
if "%BASE_PY_MODE%"=="python" (
  python %*
  exit /b %errorlevel%
)
if "%BASE_PY_MODE%"=="py" (
  py -3 %*
  exit /b %errorlevel%
)
exit /b 1