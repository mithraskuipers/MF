@echo off
setlocal
set "DOCXEX_PS1=%~dp0Extract-SrcFilesFromDocx.ps1"
if not exist "%DOCXEX_PS1%" goto :missing
rem Run the script text in memory: signing policies (AllSigned via Group Policy) only apply to script files.
powershell.exe -NoProfile -STA -Command "& ([ScriptBlock]::Create([System.IO.File]::ReadAllText($env:DOCXEX_PS1)))"
set "RC=%ERRORLEVEL%"
echo.
pause
exit /b %RC%
:missing
echo Cannot find "%DOCXEX_PS1%". Keep both files in the same folder.
pause
exit /b 1
