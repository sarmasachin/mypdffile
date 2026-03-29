@echo off
setlocal
cd /d C:\Users\DELL\Desktop\pdf-editor-tool
call stop-pdf-editor.bat

title PDF Editor Server - DO NOT CLOSE
color 0A

set "PY_EXE=C:\Python314\python.exe"
if not exist "%PY_EXE%" set "PY_EXE=python"

echo ===============================================
echo PDF Editor server running on http://127.0.0.1:8000
echo Keep this window open.
echo ===============================================

:loop
%PY_EXE% run_server.py

echo.
echo Server stopped. Auto-restarting in 2 seconds...
timeout /t 2 >nul
goto loop
