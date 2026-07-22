@echo off
chcp 65001 >nul
echo ============================================================
echo    电商审单核对 - 一键启动并生成「公网链接」
echo ============================================================
echo.
echo  用途：双击本文件，会启动网页服务并打通公网隧道，
echo        屏幕上出现的 https 链接发给同事，对方打开即可使用。
echo.
cd /d "%~dp0"

echo [1/2] 正在启动本地网页服务(Flask)...
set PORT=5000
start "审单核对后端" python app.py
timeout /t 4 >nul

echo [2/2] 正在打通公网隧道(Cloudflare)...
echo.
echo  稍候，下方会显示一行类似：
echo      Your quick Tunnel has been created at: https://xxxx.trycloudflare.com
echo  把这整串 https 链接复制发给同事即可。
echo.
echo  【重要】保持此窗口一直打开；关闭窗口 = 停止服务、链接失效。
echo.
cloudflared tunnel --url http://localhost:5000
