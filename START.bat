@echo off
title Assembly Hub Server
echo.
echo  ================================
echo   ASSEMBLY HUB - Starting...
echo  ================================
echo.
cd /d "%~dp0"
python server.py
pause
