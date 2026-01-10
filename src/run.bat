# run with ps1

rem @echo off
chcp 65001 >nul
powershell.exe -NoExit -ExecutionPolicy Bypass -NoProfile -File "%~dp0mouse_wheel_scroll_reverse.ps1"
if errorlevel 1 pause
