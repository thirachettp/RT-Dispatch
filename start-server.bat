@echo off
setlocal enabledelayedexpansion
title SCDC-MVP Server (LAN Mode)
cd /d "%~dp0"

echo ========================================
echo         SCDC-MVP SERVER (LAN MODE)
echo ========================================
echo.
echo This will start SCDC-MVP so people on the same
echo Wi-Fi / network can use it from their own browser.
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
echo.
echo   Please install Python 3.10 or newer from:
echo   https://www.python.org/downloads/
echo.
echo   IMPORTANT: during install, check the box that says
echo   "Add Python to PATH", then run this file again.
echo.
pause
exit /b 1

:have_python
rem --- Step 2: create a virtual environment if needed ---
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
echo Checking dependencies...
python -m pip install --quiet --disable-pip-version-check -r "%~dp0backend\requirements.txt"
if not errorlevel 1 goto deps_ok
echo [ERROR] Failed to install dependencies. Check your internet connection.
pause
exit /b 1

:deps_ok
rem --- Step 4: detect this computer's REAL LAN IPv4 address ---
rem Naively grabbing the first "IPv4 Address" line from ipconfig often picks
rem up a virtual adapter instead (Docker, WSL, VPN, VMware, VirtualBox, Hyper-V)
rem which is NOT reachable from other devices. This walks ipconfig adapter by
rem adapter and only trusts "Ethernet adapter" / "Wireless LAN adapter"
rem sections whose name doesn't look virtual.
echo.
echo Detecting this computer's network address...
set FIRSTIP=
set FALLBACK_IP=
set ADAPTER_OK=0

for /f "usebackq delims=" %%L in (`ipconfig`) do (
    set "LINE=%%L"
    call :scdc_check_line
)
goto detect_done

:scdc_check_line
echo !LINE! | findstr /B /C:"Ethernet adapter" /C:"Wireless LAN adapter" >nul
if errorlevel 1 goto scdc_check_ip
echo !LINE! | findstr /I /C:"vEthernet" /C:"VirtualBox" /C:"VMware" /C:"Docker" /C:"WSL" /C:"Loopback" /C:"Tailscale" /C:"Hamachi" /C:"Radmin" /C:"TAP-" /C:"Npcap" /C:"Hyper-V" /C:"Bluetooth" >nul
if not errorlevel 1 (set ADAPTER_OK=0) else (set ADAPTER_OK=1)
goto :eof

:scdc_check_ip
echo !LINE! | findstr /C:"IPv4 Address" >nul
if errorlevel 1 goto :eof
for /f "tokens=2 delims=:" %%A in ("!LINE!") do (
    set "IP=%%A"
    set "IP=!IP: =!"
    if "!ADAPTER_OK!"=="1" (
        if not defined FIRSTIP echo   found: !IP! ^(Wi-Fi/Ethernet - looks correct^)
        if not defined FIRSTIP set "FIRSTIP=!IP!"
    ) else (
        if not defined FALLBACK_IP set "FALLBACK_IP=!IP!"
    )
)
goto :eof

:detect_done
if defined FIRSTIP goto have_ip
if not defined FALLBACK_IP goto have_ip
echo   Could not confidently identify your Wi-Fi/Ethernet adapter.
echo   Falling back to !FALLBACK_IP! - this might belong to a VPN, Docker,
echo   WSL, or other virtual network and may NOT be reachable from other
echo   devices. If other devices can't connect, run "ipconfig" yourself and
echo   use the IPv4 Address listed under "Wireless LAN adapter Wi-Fi" or
echo   "Ethernet adapter Ethernet" instead.
set FIRSTIP=%FALLBACK_IP%

:have_ip
rem --- Step 5: try to open Windows Firewall (private only) ---
echo Checking Windows Firewall...
netsh advfirewall firewall show rule name="SCDC-MVP LAN Access" >nul 2>nul
if errorlevel 1 goto add_fw_rule
echo Firewall rule already exists.
goto fw_done

:add_fw_rule
netsh advfirewall firewall add rule name="SCDC-MVP LAN Access" dir=in action=allow protocol=TCP localport=8000 profile=private >nul 2>nul
netsh advfirewall firewall add rule name="SCDC-MVP LAN Access HTTPS" dir=in action=allow protocol=TCP localport=8443 profile=private >nul 2>nul
if errorlevel 1 goto fw_warn
echo Firewall rule created (private network only).
goto fw_done

:fw_warn
echo.
echo [WARNING] Could not create a Windows Firewall rule automatically.
echo   This usually means this window is not running as Administrator.
echo   Other devices may not be able to connect until you either:
echo     a^) Right-click start-server.bat and choose "Run as administrator", or
echo     b^) Manually allow TCP port 8000/8443 in Windows Firewall.
echo   The server will still start and work on this computer either way.
echo.

:fw_done
rem --- Step 6: use HTTPS if setup-https.bat has already generated a cert ---
set USE_HTTPS=0
if exist "%~dp0certs\cert.pem" if exist "%~dp0certs\key.pem" set USE_HTTPS=1

echo.
echo ========================================
echo Server is starting...
echo.
if "%USE_HTTPS%"=="1" goto show_https_info
goto show_http_info

:show_https_info
echo HTTPS is enabled - push notifications and "install as app" will work.
echo.
echo This computer:
echo   https://localhost:8443
echo.
if not defined FIRSTIP goto no_ip_https
echo Other devices on the same network ^(PC, laptop, phone^):
echo   https://!FIRSTIP!:8443
echo.
echo Note: other devices must trust this certificate first - see the
echo instructions setup-https.bat printed when you ran it.
goto show_footer

:no_ip_https
echo Could not auto-detect a LAN IP address. Run "ipconfig" and look
echo for "IPv4 Address" under your Wi-Fi/Ethernet adapter, then use that
echo with :8443 and https://
goto show_footer

:show_http_info
echo Running over plain HTTP - push notifications and PWA install will
echo NOT work. Run setup-https.bat once to enable HTTPS, then start this
echo file again.
echo.
echo This computer:
echo   http://localhost:8000
echo.
if not defined FIRSTIP goto no_ip_http
echo Other devices on the same network ^(PC, laptop, phone^):
echo   http://!FIRSTIP!:8000
echo.
echo If that address doesn't work from another device, run "ipconfig" on
echo this computer and double check the IPv4 Address under your actual
echo Wi-Fi or Ethernet adapter section (not VPN/Docker/WSL/Virtual ones).
goto show_footer

:no_ip_http
echo Could not auto-detect a LAN IP address. Run "ipconfig" and look
echo for "IPv4 Address" under your Wi-Fi/Ethernet adapter, then use that
echo with :8000

:show_footer
echo.
echo Press CTRL+C in this window to stop the server.
echo ========================================
echo.

cd /d "%~dp0backend"
if "%USE_HTTPS%"=="1" goto launch_https
goto launch_http

:launch_https
python -m uvicorn main:app --host 0.0.0.0 --port 8443 --ssl-keyfile="%~dp0certs\key.pem" --ssl-certfile="%~dp0certs\cert.pem"
goto server_stopped

:launch_http
python -m uvicorn main:app --host 0.0.0.0 --port 8000

:server_stopped
echo.
echo Server stopped.
pause
