@echo off
:: Markus Scholtes, 2023
:: Compile VirtualDesktop in .Net 4.x environment
setlocal

echo %CMDCMDLINE%

set versions=VirtualDesktop11,VirtualDesktop11-23H2,VirtualDesktop11InsiderCanary,VirtualDesktop11-21H2,VirtualDesktop,VirtualDesktopServer2022,VirtualDesktopServer2016

del .\v*.exe 2>nul

for %%i in (%versions%) do (
     echo Compiling %%i.cs ...
     C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe "%~dp0%%i.cs" /win32icon:"%~dp0MScholtes.ico" 2> nul > nul
     ren "%~dp0%%i.exe"  vd.exe
     echo Testing %%i.exe  ...
     "%~dp0vd.exe" /LIST 2>nul >nul
     if ERRORLEVEL 0 ( echo compiled %%i.cs to VirtualDesktop.exe )
     if ERRORLEVEL 0 ( goto done ) else ( del "%~dp0vd.exe" )
    
)

:done

if exist .\vd.exe ren .\vd.exe VirtualDesktop.exe
 