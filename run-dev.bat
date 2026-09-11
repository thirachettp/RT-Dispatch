@echo off
setlocal enabledelayedexpansion
title SCDC-MVP Server (Dev Mode)
cd /d "%~dp0"

echo ========================================
echo       SCDC-MVP SERVER (DEV MODE)
echo ========================================
echo.
echo Local development only - auto-reloads on code changes.
echo Not reachable from other devices. For LAN/team use,
echo run start-server.bat instead.
echo.

rem --- Step 1: find Python (goto-based, avoids parenthesis-in-path issues) ---
set PYEXE=
where python >nul 2>nul
if not errorlevel 1 set PYEXE=python
if defined PYEXE goto have_python
where py >nul 2>nul
if not errorlevel 1 set PYEXE=py
if defined PYEXE goto have_python

echo [ERROR] Python was not found on this computer.
echo   Install it from https://www.python.org/downloads/
echo   and check "Add Python to PATH" during setup.
pause
exit /b 1

:have_python
rem --- Step 2: create/activate the virtual environment ---
set VENV_DIR=%~dp0.venv
if exist "%VENV_DIR%\Scripts\python.exe" goto venv_ready
echo Setting up the app for the first time, please wait...
%PYEXE% -m venv "%VENV_DIR%"
if exist "%VENV_DIR%\Scripts\python.exe" goto venv_ready
echo [ERROR] Could not create the virtual environment.
pause
exit /b 1

:venv_ready
call "%VENV_DIR%\Scripts\activate.bat"

rem --- Step 3: install/update dependencies ---
python -m pip install --quiet --disable-pip-version-check -r "%~dp0backend\requirements.txt"
if not errorlevel 1 goto deps_ok
echo [ERROR] Failed to install dependencies. Check your internet connection.
pause
exit /b 1

:deps_ok
rem --- Step 4: start the dev server (localhost only, auto-reload) ---
echo.
echo Open http://localhost:8000 in your browser.
echo Press CTRL+C to stop.
echo.

cd /d "%~dp0backend"
python -m uvicorn main:app --reload

echo.
echo Server stopped.
pause
