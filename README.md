# VirtualDesktop
V1.4.1, 2019-07-15

C# command line tool to manage virtual desktops in Windows 10<br><br>
(look for a powershell version here: https://gallery.technet.microsoft.com/Powershell-commands-to-d0e79cc5)

**With Windows 10 1809 Microsoft changed the API (COM GUIDs) for accessing the functions for virtual desktops again. I provide three versions of virtualdesktop.cs now: virtualdesktop.cs is for Windows 10 1809 and newer, virtualdesktop1803.cs is for Windows 10 1803, virtualdesktop1607.cs is for Windows 10 1607 to 1709 and Windows Server 2016. Using Compile.bat all executables  will be generated (thanks to [mzomparelli](https://github.com/mzomparelli/zVirtualDesktop/wiki) for investigating).**

## Generate:
Compile with Compile.bat (no visual studio needed, but obviously Windows 10)

## Description:
Command line tool to manage the virtual desktops of Windows 10.
Parameters can be given as a sequence of commands. The result - most of the times the number of the processed desktop - can be used as input for the next parameter. The result of the last command is returned as error level.
Virtual desktop numbers start with 0.

## Parameters (leading / can be omitted or - can be used instead):
**/Help /h /?**      this help screen.

**/Verbose /Quiet**  enable verbose (default) or quiet mode (short: /v and /q).

**/Break /Continue** break (default) or continue on error (short: /b and /co).

**/Count**           get count of virtual desktops to pipeline (short: /c).

**/GetDesktop:(n)**  get number of virtual desktop (n) to pipeline (short: /gd).

**/GetCurrentDesktop**  get number of current desktop to pipeline (short: /gcd).

**/IsVisible[:(n)]**  is desktop number (n) or number in pipeline visible (short: /iv)? Returns 0 for visible and 1 for invisible.

**/Switch[:(n)]**    switch to desktop with number (n) or with number in pipeline (short: /s).

**/Left**            switch to virtual desktop to the left of the active desktop (short: /l).

**/Right**           switch to virtual desktop to the right of the active desktop (short: /ri).

**/New**             create new desktop (short: /n). Number is stored in pipeline.

**/Remove[:(n)]**    remove desktop number (n) or desktop with number in pipeline (short: /r).

**/MoveWindow:(s)**  move process with name (s) to desktop with number in pipeline (short: /mw).

**/MoveWindow:(n)**  move process with id (n) to desktop with number in pipeline (short: /mw).

**/MoveWindowHandle:(n)**  move window with handle (n) to desktop with number in pipeline (short: /mwh).

**/MoveActiveWindow**  move active window to desktop with number in pipeline (short: /maw).

**/GetDesktopFromWindow:(s)**  get desktop number where process with name (s) is displayed (short: /gdfw).

**/GetDesktopFromWindow:(n)**  get desktop number where process with id (n) is displayed (short: /gdfw).

**/IsWindowOnDesktop:(s)**  check if process with name (s) is on desktop with number in pipeline (short: /iwod). Returns 0 for yes, 1 for no.

**/IsWindowOnDesktop:(n)**  check if process with id (n) is on desktop with number in pipeline (short: /iwod). Returns 0 for yes, 1 for no.

**/PinWindow:(s)**   pin process with name (s) to all desktops (short: /pw).

**/PinWindow:(n)**   pin process with id (n) to all desktops (short: /pw).

**/UnPinWindow:(s)**  unpin process with name (s) from all desktops (short: /upw).

**/UnPinWindow:(n)**  unpin process with id (n) from all desktops (short: /upw).

**/IsWindowPinned:(s)**  check if process with name (s) is pinned to all desktops (short: /iwp). Returns 0 for yes, 1 for no.

**/IsWindowPinned:(n)**  check if process with id (n) is pinned to all desktops (short: /iwp). Returns 0 for yes, 1 for no.

**/PinApplication:(s)**  pin application with name (s) to all desktops (short: /pa).

**/PinApplication:(n)**  pin application with process id (n) to all desktops (short: /pa).

**/UnPinApplication:(s)**  unpin application with name (s) from all desktops (short: /upa).

**/UnPinApplication:(n)**  unpin application with process id (n) from all desktops (short: /upa).

**/IsApplicationPinned:(s)**  check if application with name (s) is pinned to all desktops (short: /iap). Returns 0 for yes, 1 for no.

**/IsApplicationPinned:(n)**  check if application with process id (n) is pinned to all desktops (short: /iap). Returns 0 for yes, 1 for no.

**/WaitKey**       wait for key press (short: /wk).

**/Sleep:(n)**     wait for (n) milliseconds (short: /sl).

## Examples:
```
Virtualdesktop.exe -New -Switch -GetCurrentDesktop
Virtualdesktop.exe sleep:200 gd:1 mw:notepad s
Virtualdesktop.exe /Count /continue /Remove /Remove /Count
VirtualDesktop.exe -IsWindowPinned:cmd
if ERRORLEVEL 1 VirtualDesktop.exe PinWindow:cmd
```
