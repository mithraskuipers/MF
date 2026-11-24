@echo off
setlocal
set "SCRIPT=%~dp0Find-CobolSql.ps1"
if not exist "%SCRIPT%" goto :missing
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%SCRIPT%"
set "RC=%ERRORLEVEL%"
echo.
pause
exit /b %RC%
:missing
echo Cannot find "%SCRIPT%". Keep both files in the same folder.
pause
exit /b 1
