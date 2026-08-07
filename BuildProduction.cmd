@echo off
setlocal
cd /d "%~dp0"

echo MailNotes Add-in - Production Build
echo.
echo Die Build-Einstellungen stehen zentral in mailnotes.config.js.
echo.
call npm run build
if errorlevel 1 exit /b %errorlevel%

echo.
echo Fertig: dist\
endlocal
