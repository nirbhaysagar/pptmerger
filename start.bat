@echo off
title PPT Unify Launcher
echo ========================================================
echo        Starting PPT Unify Presentation Merger
echo ========================================================
cd /d "%~dp0"

REM Check for virtual environment
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" run.py %*
    goto end
)

REM Check for python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed or not in PATH!
    echo Please install Python 3 from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation.
    pause
    exit /b 1
)

REM Run launcher
python run.py %*

:end
