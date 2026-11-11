@echo off
title Uninstall Deli Elements for PowerPoint
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "6f1d2b7e-3c4a-4d8e-9a5b-de11e1e2027a" /f >nul 2>&1
if exist "%LOCALAPPDATA%\DeliElements" rd /s /q "%LOCALAPPDATA%\DeliElements"
if exist "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" rd /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" 2>nul
echo.
echo  Deli Elements has been removed. Restart PowerPoint to finish.
echo.
pause
