@echo off
setlocal
title Install Deli Elements for PowerPoint
echo.
echo  ===============================================
echo   Deli Elements for PowerPoint - installer
echo  ===============================================
echo.

tasklist /FI "IMAGENAME eq POWERPNT.EXE" 2>nul | find /I "POWERPNT.EXE" >nul
if not errorlevel 1 (
  echo  PowerPoint is open. Please save your work and close PowerPoint,
  echo  then press any key to continue.
  pause >nul
)

set "DIR=%LOCALAPPDATA%\DeliElements"
if not exist "%DIR%" mkdir "%DIR%"

echo  Downloading the add-in...
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { [Net.ServicePointManager]::SecurityProtocol = 'Tls12'; Invoke-WebRequest -UseBasicParsing 'https://deddytann.github.io/deli-elements/manifest.xml' -OutFile '%DIR%\manifest.xml'; exit 0 } catch { exit 1 }"
if errorlevel 1 (
  echo.
  echo  Could not download the add-in. Check your internet connection and try again.
  pause
  exit /b 1
)

echo  Registering it with PowerPoint...
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "6f1d2b7e-3c4a-4d8e-9a5b-de11e1e2027a" /t REG_SZ /d "%DIR%\manifest.xml" /f >nul
if errorlevel 1 (
  echo  Could not register the add-in.
  pause
  exit /b 1
)

rem Clear PowerPoint's add-in cache so the newest version loads
if exist "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" rd /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" 2>nul

echo.
echo  Done! Open PowerPoint and look for the "Deli" tab on the ribbon.
echo  If it isn't there: Home ^> Add-ins ^> More Add-ins ^> My Add-ins,
echo  then choose Deli Elements under Developer Add-ins.
echo.
pause
