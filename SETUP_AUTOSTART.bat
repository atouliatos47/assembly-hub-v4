@echo off
title Assembly Hub - Auto Start Setup
echo.
echo  ============================================
echo   Setting up Assembly Hub to auto-start...
echo  ============================================
echo.

set SCRIPT_DIR=%~dp0
set PYTHON_CMD=python
set TASK_NAME=AssemblyHub

:: Remove old task if exists
schtasks /delete /tn "%TASK_NAME%" /f >nul 2>&1

:: Create task to run on login
schtasks /create /tn "%TASK_NAME%" ^
  /tr "\"%PYTHON_CMD%\" \"%SCRIPT_DIR%server.py\"" ^
  /sc ONLOGON ^
  /delay 0000:30 ^
  /rl HIGHEST ^
  /f

if %errorlevel% == 0 (
    echo.
    echo  SUCCESS! Assembly Hub will now start automatically
    echo  when you log into Windows.
    echo.
    echo  To remove auto-start, run REMOVE_AUTOSTART.bat
) else (
    echo.
    echo  Could not create task. Try running as Administrator.
)
echo.
pause
