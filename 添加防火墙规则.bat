@echo off
chcp 65001 >nul
echo ========================================
echo   添加防火墙规则 - 允许5000端口访问
echo ========================================
echo.

:: 检查是否有管理员权限
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 需要以管理员身份运行！
    echo.
    echo 请右键点击此文件，选择"以管理员身份运行"
    pause
    exit /b 1
)

echo 正在添加防火墙规则...
netsh advfirewall firewall delete rule name="Flask_WebApp_5000" >nul 2>&1
netsh advfirewall firewall add rule name="Flask_WebApp_5000" dir=in action=allow protocol=TCP localport=5000

if %errorlevel% equ 0 (
    echo.
    echo [成功] 防火墙规则已添加！
    echo 现在其他电脑可以通过 http://172.16.10.84:5000 访问了
) else (
    echo.
    echo [失败] 添加防火墙规则失败，请手动添加
)

echo.
pause
