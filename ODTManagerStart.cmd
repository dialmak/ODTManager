@echo off
color 1f
title Clean Licenses Office 
setlocal enableextensions

call :CheckIsAdmin || goto :Exit
cd /d "%~dp0"
if defined PROCESSOR_ARCHITEW6432 start "" %SystemRoot%\sysnative\cmd.exe /c "%~nx0" %* & goto :EOF
set PATH=%cd%;%SystemRoot%\system32;%SystemRoot%;%SystemRoot%\System32\Wbem;%SystemRoot%\System32\WindowsPowerShell\v1.0\;
reg add "HKCU\Software\Microsoft\Windows Script Host\Settings" /v "Enabled" /t REG_DWORD /d "1" /f >nul 2>nul
reg add "HKLM\Software\Microsoft\Windows Script Host\Settings" /v "Enabled" /t REG_DWORD /d "1" /f >nul 2>nul
set "xOS=x64"
if /i "%PROCESSOR_ARCHITECTURE%"=="x86" if not defined PROCESSOR_ARCHITEW6432 set "xOS=x86"
cd /d "%cd%\bin\%xOS%"
cleanospp >nul 2>nul

:Exit
endlocal
echo.
rem echo Press any key to exit ...
rem pause>nul
goto :eof

:CheckIsAdmin
set guid=%random%%random%-%random%-%random%%random%%random%
mkdir %windir%\%guid%>nul 2>&1
rmdir %windir%\%guid%>nul 2>&1 || (
  echo Error. Admin privileges required.
  exit /b 1
)
exit /b 0