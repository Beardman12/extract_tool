@echo off
setlocal EnableExtensions
chcp 65001 >nul

rem 切到脚本所在目录（仓库根目录）
cd /d "%~dp0"

echo [1/5] 检查 Python 虚拟环境...
if not exist ".venv\Scripts\python.exe" (
  echo 未检测到 .venv，开始创建...
  py -3.14 -m venv .venv >nul 2>&1
  if errorlevel 1 (
    py -3 -m venv .venv
    if errorlevel 1 (
      echo 创建虚拟环境失败，请确认已安装 Python 3.14+ 或 py 启动器可用。
      pause
      exit /b 1
    )
  )
)

set "PYTHON=%CD%\.venv\Scripts\python.exe"

echo [2/5] 准备 pip...
"%PYTHON%" -m ensurepip --upgrade >nul 2>&1

echo [3/5] 安装/更新依赖...
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

echo [4/5] 检查 .env 邮箱配置...
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

set "API_PORT=8003"
set /p API_PORT=请输入 API 端口（默认 8003）:
if "%API_PORT%"=="" set "API_PORT=8003"

echo [5/5] 启动 API 和定时服务...
start "AI Price API" cmd /k ""%PYTHON%" scripts\run_api.py --port %API_PORT%"
start "AI Price Scheduler" cmd /k ""%PYTHON%" scripts\run_scheduler.py --interval 300"

echo.
echo 启动完成：
echo API 地址: http://127.0.0.1:%API_PORT%/docs
echo Scheduler 间隔: 300 秒
echo.
echo 关闭服务：直接关闭对应命令行窗口即可。
pause