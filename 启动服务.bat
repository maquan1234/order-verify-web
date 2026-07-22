@echo off
chcp 65001 >nul
title 电商审单核对系统 - 公网版
echo ========================================
echo   电商审单核对系统 - 启动中...
echo   (自动创建公网隧道，任何网络可访问)
echo ========================================
echo.

cd /d "C:\Users\EDY\WorkBuddy\2026-07-21-16-17-50\webapp"

"C:\Users\EDY\.workbuddy\binaries\python\versions\3.12.8\python.exe" start.py

pause
