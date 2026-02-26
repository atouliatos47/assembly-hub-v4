@echo off
set STARTUP=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup
del "%STARTUP%\AssemblyHub.vbs" >nul 2>&1
echo Assembly Hub removed from startup.
pause
