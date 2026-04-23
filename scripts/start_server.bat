@echo off
set "PORT=3005"
echo Current Version: 1.2.0 (Advanced Auth)
echo Starting MPS Dashboard Server on Port %PORT%...
cd /d "%~dp0.."
node server/server.js
pause
