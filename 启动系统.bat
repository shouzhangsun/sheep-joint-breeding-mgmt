@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 启动山羊育种管理平台...
C:\Users\Administrator\.workbuddy\binaries\python\versions\3.13.12\python.exe app.py
pause
