@echo off
title Assembly Hub - Remove Auto Start
schtasks /delete /tn "AssemblyHub" /f
echo Assembly Hub auto-start removed.
pause
