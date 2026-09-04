@echo off
REM ============================================================
REM  山羊联合育种管理系统 — 一键构建（Windows）
REM  前置：Python(含 flask/openpyxl/pywebview/reportlab/pyinstaller) + Inno Setup 6
REM  步骤：生成空库 -> PyInstaller 打包 -> 自签名并签名 -> 编译安装包
REM  说明：
REM   - 若项目根目录存在 venv\ 或 .venv\，则自动优先使用其中的 python/pyinstaller；
REM     否则使用 PATH 中激活环境里的 python/pyinstaller（请先 activate 你的虚拟环境）。
REM   - ISCC 若不在 PATH，则自动探测常见 Inno Setup 6 安装位置。
REM   - 自签名证书每次构建都会重新生成（make_cert_and_sign.ps1），
REM     安装包会在目标机以管理员权限把证书导入「受信任根证书颁发机构」以消除 SmartScreen。
REM ============================================================
setlocal
cd /d %~dp0

REM ---- 自动定位 Python / PyInstaller ----
REM 重要：本项目打包为 32 位 exe（兼容 x86 WebView2 运行时，最通用），
REM 必须使用 32 位 venv（win32）里的 python/pyinstaller，不能用 64 位。
set PYTHON=python
set PYINSTALLER=pyinstaller
if exist "%~dp0venv\Scripts\python.exe" (
  set "PYTHON=%~dp0venv\Scripts\python.exe"
  set "PYINSTALLER=%~dp0venv\Scripts\pyinstaller.exe"
) else if exist "%~dp0.venv\Scripts\python.exe" (
  set "PYTHON=%~dp0.venv\Scripts\python.exe"
  set "PYINSTALLER=%~dp0.venv\Scripts\pyinstaller.exe"
) else if exist "C:\Users\Administrator\.workbuddy\binaries\python\envs\win32\Scripts\python.exe" (
  set "PYTHON=C:\Users\Administrator\.workbuddy\binaries\python\envs\win32\Scripts\python.exe"
  set "PYINSTALLER=C:\Users\Administrator\.workbuddy\binaries\python\envs\win32\Scripts\pyinstaller.exe"
)

REM ---- 自动定位 ISCC（Inno Setup 编译器）----
set ISCC=iscc
where iscc >nul 2>nul
if errorlevel 1 (
  if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" set ISCC="C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
  if exist "C:\Program Files\Inno Setup 6\ISCC.exe" set ISCC="C:\Program Files\Inno Setup 6\ISCC.exe"
  if exist "%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe" set ISCC="%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe"
  if exist "%APPDATA%\Programs\Inno Setup 6\ISCC.exe" set ISCC="%APPDATA%\Programs\Inno Setup 6\ISCC.exe"
)

echo [1/4] 生成内置空库（6 品种）
%PYTHON% make_seed.py
if errorlevel 1 goto :fail

echo [2/4] PyInstaller 打包（onefile 单文件）
%PYINSTALLER% build.spec --noconfirm
if errorlevel 1 goto :fail

echo [3/4] 生成自签名证书并签名 exe
powershell -ExecutionPolicy Bypass -File make_cert_and_sign.ps1
if errorlevel 1 goto :fail

echo [4/4] 编译 Inno Setup 安装包
%ISCC% installer.iss
if errorlevel 1 goto :fail

echo.
echo ========== 构建完成 ==========
echo 安装包位于 installer\Setup_山羊联合育种管理系统.exe
goto :end

:fail
echo.
echo [构建失败] 请检查上方错误输出。
pause
exit /b 1

:end
pause
