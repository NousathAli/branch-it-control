@echo off
cd /d "%~dp0"
echo Starting Branch IT Control Dashboard v8.23 Store Print A4 Readable...
echo.
echo Local: http://localhost:3000
echo Tailscale/Mobile: http://YOUR-TAILSCALE-IP:3000
echo Status: http://localhost:3000/api/status
echo.
echo Shared database file: data\db.json
echo.
node server.js
pause
