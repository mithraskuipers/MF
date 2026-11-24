@echo off
setlocal
set "CBLSQL_PS1=%~dp0Find-CobolSql.ps1"
if not exist "%CBLSQL_PS1%" goto :missing
rem Run the script text in memory: signing policies (AllSigned via Group Policy) only apply to script files.
powershell.exe -NoProfile -STA -Command "& ([ScriptBlock]::Create([System.IO.File]::ReadAllText($env:CBLSQL_PS1)))"
set "RC=%ERRORLEVEL%"
echo.
pause
exit /b %RC%
:missing
echo Cannot find "%CBLSQL_PS1%". Keep both files in the same folder.
pause
exit /b 1
