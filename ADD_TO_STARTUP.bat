@echo off
title Assembly Hub - Add to Startup
echo.
echo  Adding Assembly Hub to Windows Startup folder...
echo.

set "STARTUP=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
set "SCRIPT_DIR=%~dp0"
set "VBS=%STARTUP%\AssemblyHub.vbs"

:: Write the VBS launcher
(
echo Set oShell = CreateObject^("WScript.Shell"^)
echo oShell.Run "python ""%SCRIPT_DIR%server.py""", 0, False
) > "%VBS%"

if exist "%VBS%" (
    echo.
    echo  SUCCESS! Assembly Hub will now start automatically
    echo  on login with NO console window.
    echo.
    echo  Location: %VBS%
    echo.
    echo  To remove: run REMOVE_STARTUP.bat
) else (
    echo  FAILED - could not write to startup folder.
    echo  Path: %STARTUP%
)
echo.
pause
