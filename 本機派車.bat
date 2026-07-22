@echo off
chcp 65001 >nul
set "HERE=%~dp0"
cd /d "%HERE%"

set "PY="
where py >nul 2>&1 && set "PY=py"
if not defined PY where python >nul 2>&1 && set "PY=python"
if not defined PY if exist "%LOCALAPPDATA%\hermes\hermes-agentenv\Scripts\python.exe" set "PY=%LOCALAPPDATA%\hermes\hermes-agentenv\Scripts\python.exe"
if not defined PY (
    echo Python not found. Install Python with py launcher, or check Hermes install.
    pause
    exit /b 1
)

echo Running local route planning, please wait...
%PY% plan_and_push_line.py
echo:
echo Done. See report path printed above.
pause
