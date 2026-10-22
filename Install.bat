@echo off
chcp 65001 >nul
title Partnership Tools - Install
cd /d "%~dp0"
if not exist "%~dp0app\pyembed\python.exe" goto missing
"%~dp0app\pyembed\python.exe" "%~dp0app\install.py"
if errorlevel 1 (pause & exit /b 1)
timeout /t 8 >nul
exit /b

:missing
echo Missing files in the app folder. Extract the WHOLE zip first, then run Install.bat.
pause
