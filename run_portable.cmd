@echo off
setlocal enableextensions
set SCRIPT_DIR=%~dp0
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%run_portable.ps1"
endlocal
