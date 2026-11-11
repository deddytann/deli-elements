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

echo  Setting The Deli Group template as the default for new presentations...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; try { [Net.ServicePointManager]::SecurityProtocol='Tls12'; $base='https://deddytann.github.io/deli-elements/templates'; $t=Join-Path $env:APPDATA 'Microsoft\Templates'; New-Item -ItemType Directory -Force -Path $t | Out-Null; $b=Join-Path $t 'Blank.potx'; $bk=Join-Path $t 'Blank.before-deli.potx'; if ((Test-Path $b) -and -not (Test-Path $bk)) { Copy-Item $b $bk }; Invoke-WebRequest -UseBasicParsing ($base+'/Blank.potx') -OutFile $b; $k='HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Options'; if (-not (Test-Path $k)) { New-Item -Path $k | Out-Null }; $p=(Get-ItemProperty -Path $k -Name PersonalTemplates -ErrorAction SilentlyContinue).PersonalTemplates; if (-not $p) { $p=Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'Custom Office Templates'; New-ItemProperty -Path $k -Name PersonalTemplates -PropertyType ExpandString -Value $p -Force | Out-Null }; $p=[Environment]::ExpandEnvironmentVariables($p); New-Item -ItemType Directory -Force -Path $p | Out-Null; Invoke-WebRequest -UseBasicParsing ($base+'/The_Deli_Group_Template.potx') -OutFile (Join-Path $p 'The Deli Group.potx'); exit 0 } catch { Write-Host $_; exit 1 }"
if errorlevel 1 echo  Could not install the template (the add-in itself is installed).

rem Clear PowerPoint's add-in cache so the newest version loads
if exist "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" rd /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" 2>nul

echo.
echo  Done! Open PowerPoint: new presentations now start on The Deli Group template,
echo  the full template is under New ^> Personal, and the "Deli" tab is on the ribbon.
echo  If it isn't there: Home ^> Add-ins ^> More Add-ins ^> My Add-ins,
echo  then choose Deli Elements under Developer Add-ins.
echo.
pause
