@echo off
title Uninstall Deli Elements for PowerPoint
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "6f1d2b7e-3c4a-4d8e-9a5b-de11e1e2027a" /f >nul 2>&1
if exist "%LOCALAPPDATA%\DeliElements" rd /s /q "%LOCALAPPDATA%\DeliElements"
if exist "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" rd /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" 2>nul
powershell -NoProfile -ExecutionPolicy Bypass -Command "$t=Join-Path $env:APPDATA 'Microsoft\Templates'; $b=Join-Path $t 'Blank.potx'; $bk=Join-Path $t 'Blank.before-deli.potx'; if (Test-Path $bk) { Move-Item -Force $bk $b } elseif (Test-Path $b) { Remove-Item $b }; $p=(Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Options' -Name PersonalTemplates -ErrorAction SilentlyContinue).PersonalTemplates; if ($p) { $f=Join-Path ([Environment]::ExpandEnvironmentVariables($p)) 'The Deli Group.potx'; if (Test-Path $f) { Remove-Item $f } }"
echo.
echo  Deli Elements and the Deli default template have been removed. Restart PowerPoint to finish.
echo.
pause
