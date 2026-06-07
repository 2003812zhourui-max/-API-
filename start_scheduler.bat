@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ========================================
echo   领星自动化调度平台
echo ========================================
echo.
echo 启动中, 浏览器将自动打开...
echo 关闭此窗口即停止调度服务。
echo.
start http://localhost:5000
"C:\Users\Administrator\AppData\Local\Programs\Python\Python314\python.exe" -m scheduler_app.app
pause
