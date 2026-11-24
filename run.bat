@echo off
cd /d %~dp0
set PORTABLE_PY=%~dp0python\python.exe
set RUN_PY=python
if exist "%PORTABLE_PY%" set RUN_PY="%PORTABLE_PY%"

echo ========================================================
echo   Project Cost Local App v4.0 Starting...
echo   URL: http://127.0.0.1:8765
echo ========================================================
cd /d "%~dp0backend"
start "CostAppServer_v4" cmd /k %RUN_PY% -m uvicorn app:app --host 127.0.0.1 --port 8765
timeout /t 2 /nobreak > nul
start "" "http://127.0.0.1:8765"
