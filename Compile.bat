@echo off
:: Markus Scholtes, 2025
:: Compile VirtualDesktop in .Net 4.x environment
setlocal

C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe "%~dp0VirtualDesktop11.cs" /win32icon:"%~dp0VirtualDesktop.ico"
C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe "%~dp0VirtualDesktop11-24H2.cs" /win32icon:"%~dp0VirtualDesktop.ico"
C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe "%~dp0VirtualDesktop.cs" /win32icon:"%~dp0VirtualDesktop.ico"
C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe "%~dp0VirtualDesktopServer2022.cs" /win32icon:"%~dp0VirtualDesktop.ico"
C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe "%~dp0VirtualDesktopServer2016.cs" /win32icon:"%~dp0VirtualDesktop.ico"

:: was batch started in Windows Explorer? Yes, then pause
echo "%CMDCMDLINE%" | find /i "/c" > nul
if %ERRORLEVEL%==0 pause
