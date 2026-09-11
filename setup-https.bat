@echo off
setlocal enabledelayedexpansion
title SCDC-MVP HTTPS Setup (mkcert)
cd /d "%~dp0"

echo ========================================
echo    SCDC-MVP HTTPS SETUP (mkcert)
echo ========================================
echo.
echo This creates a locally-trusted HTTPS certificate so browsers
echo allow Web Push notifications and PWA install on your LAN.
echo Run this once (and again if this computer's IP address changes).
echo.

rem --- Step 1: check mkcert is installed ---
where mkcert >nul 2>nul
if not errorlevel 1 goto have_mkcert

echo [ERROR] mkcert was not found on this computer.
echo.
echo   Install it with ONE of these, then run this file again:
echo.
echo     winget install FiloSottile.mkcert -e
echo     choco install mkcert
echo.
echo   Or download it manually from:
echo     https://github.com/FiloSottile/mkcert/releases
echo   and put mkcert.exe somewhere on your PATH.
echo.
pause
exit /b 1

:have_mkcert
rem --- Step 2: install the local CA into this computer's trust store ---
echo Installing local certificate authority (this computer only)...
mkcert -install
if errorlevel 1 (
    echo [WARNING] "mkcert -install" reported a problem. Certificates will
    echo   still be created, but this computer's browsers may not trust
    echo   them yet. Try running this file as Administrator if that happens.
)

rem --- Step 3: detect this computer's LAN IPv4 address ---
echo.
echo Detecting this computer's network address...
set FIRSTIP=
for /f "tokens=2 delims=:" %%A in ('ipconfig ^| findstr /R /C:"IPv4 Address"') do (
    set IP=%%A
    set IP=!IP: =!
    if not defined FIRSTIP set FIRSTIP=!IP!
)
if not defined FIRSTIP (
    echo [ERROR] Could not detect a LAN IP address automatically.
    echo   Run "ipconfig" yourself, find the "IPv4 Address" line, and
    echo   edit this file to add it manually if needed.
    pause
    exit /b 1
)
echo Found: !FIRSTIP!

rem --- Step 4: generate the certificate for localhost + this IP ---
if not exist "%~dp0certs" mkdir "%~dp0certs"
echo.
echo Generating certificate for localhost, 127.0.0.1, and !FIRSTIP! ...
mkcert -cert-file "%~dp0certs\cert.pem" -key-file "%~dp0certs\key.pem" localhost 127.0.0.1 !FIRSTIP!
if errorlevel 1 (
    echo [ERROR] Certificate generation failed.
    pause
    exit /b 1
)

echo.
echo ========================================
echo Certificate created successfully.
echo ========================================
echo.
echo Next: run start-server.bat as usual - it will now use HTTPS
echo automatically and print an https:// URL instead of http://.
echo.
echo IMPORTANT - for OTHER devices (phones, other PCs) to trust this
echo certificate too (needed for push notifications to work on them),
echo copy the file below to each device and install/trust it as a
echo root certificate:
echo.
for /f "delims=" %%R in ('mkcert -CAROOT') do echo   %%R\rootCA.pem
echo.
echo   Windows:  double-click it, "Install Certificate", Local Machine,
echo             "Trusted Root Certification Authorities"
echo   Android:  Settings - Security - Install a certificate - CA certificate
echo   iOS:      AirDrop/email the file, install it, THEN also enable it
echo             under Settings - General - About - Certificate Trust Settings
echo   Mac:      double-click, open Keychain Access, set it to "Always Trust"
echo.
echo If you skip this step, this server computer itself will work fine,
echo but OTHER devices will show a security warning and push notifications
echo will not work on them until they trust this certificate too.
echo ========================================
echo.
pause
