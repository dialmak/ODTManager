@echo off
color 1f
title Activation Volume Office
setlocal enableextensions
cd /d "%~dp0"
if defined PROCESSOR_ARCHITEW6432 start "" %SystemRoot%\sysnative\cmd.exe /c "%~nx0" %* & goto :EOF
set PATH=%cd%;%SystemRoot%\system32;%SystemRoot%;%SystemRoot%\System32\Wbem;%SystemRoot%\System32\WindowsPowerShell\v1.0\;
reg add "HKCU\Software\Microsoft\Windows Script Host\Settings" /v "Enabled" /t REG_DWORD /d "1" /f >nul 2>nul
call :CheckIsAdmin || goto :Exit

ping -n 2 google.com >nul 2>nul || (echo No internet connection, skip activation. & goto :Exit)
if exist "%SystemDrive%\Program Files\Microsoft Office\Office16\OSPP.VBS" (set "OsppPath=%SystemDrive%\Program Files\Microsoft Office\Office16")
if exist "%SystemDrive%\Program Files (x86)\Microsoft Office\Office16\OSPP.VBS" (set "OsppPath=%SystemDrive%\Program Files (x86)\Microsoft Office\Office16")
if not defined OsppPath (echo Error. Not found script OSPP.VBS. & goto :Exit)
cscript //nologo %windir%\system32\slmgr.vbs /skms kms.loli.beer >nul
cscript //nologo "%OsppPath%\OSPP.VBS" /sethst:kms.loli.beer | find /i /v "Processing" | find /i /v "Exiting"
cscript //nologo "%OsppPath%\OSPP.VBS" /act | find /i /v "Processing" | find /i /v "Exiting"

:Exit
endlocal
echo.
echo Press any key to exit ...
pause>nul
goto :eof

:CheckIsAdmin
set guid=%random%%random%-%random%-%random%%random%%random%
mkdir %windir%\%guid%>nul 2>&1
rmdir %windir%\%guid%>nul 2>&1 || (
  echo Error. Admin privileges required.
  exit /b 1
)
exit /b 0