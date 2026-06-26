@echo off
chcp 65001 >nul
cd /d "%~dp0"
if exist server.log del server.log
echo Starting local preview server on http://127.0.0.1:8080/
start "S低L3 HTTP Server" /min powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0server.ps1" -Port 8080 -Root "%~dp0"
timeout /t 3 >nul
echo.
echo === server.log ===
type server.log 2>nul
echo.
echo === port check ===
netstat -ano | findstr ":8080"
echo.
echo Server started. Open http://127.0.0.1:8080/ in browser.
echo To stop, run stop_server.bat
pause
