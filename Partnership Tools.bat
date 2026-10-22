@echo off
chcp 65001 >nul
title Partnership Tools - PVI Digital
cd /d "%~dp0"
if not exist "%~dp0app\pyembed\pythonw.exe" goto missing
start "" "%~dp0app\pyembed\pythonw.exe" "app\server.py"
exit /b

:missing
echo Missing files in the app folder. Extract the WHOLE zip again (to a short path, e.g. C:\Partnership-Tools), then run this file.
pause
