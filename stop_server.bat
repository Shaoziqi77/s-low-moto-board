@echo off
echo Stopping all PowerShell instances running server.ps1 ...
taskkill /F /FI "WINDOWTITLE eq S低L3 HTTP Server*" 2>nul
taskkill /F /FI "MODULES eq System.Net" 2>nul
echo Done. Port 8080 should now be free.
pause
