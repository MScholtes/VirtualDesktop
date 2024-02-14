https://github.com/MScholtes/VirtualDesktop/


VirtualDesktop

C# command line tool to manage virtual desktops in Windows 10 and Windows 11

Version 1.17, 2024-02-14
- version for Windows 11 build 22631.3085 and up

(look for a powershell version here:
https://gallery.technet.microsoft.com/Powershell-commands-to-d0e79cc5
or here:
https://www.powershellgallery.com/packages/VirtualDesktop)


With Windows 11 23H2 Release 3085 Microsoft did change the API (COM GUIDs) for 
accessing the functions for virtual desktops again. I provide five versions of 
virtualdesktop.cs now: virtualdesktop.cs is for Windows 10, virtualdesktop11.cs 
is for Windows 11 23H2 and Insider, virtualdesktop11-22h2.cs for Windows 11 22H2 
release 2215 and newer, virtualdesktopserver2022.cs is for Windows Server 2022, 
virtualdesktopserver2016.cs is for Windows Server 2016. Using Compile.bat all 
executables will be generated.

Generate:
 Compile with Compile.bat (no visual studio needed, but obviously Windows 10 or 11)

Description:
 Command line tool to manage the virtual desktops of Windows 10 and 11.
 Parameters can be given as a sequence of commands. The result - most of
 thetimes the number of the processed desktop - can be used as input for the
 next parameter. The result of the last command is returned as error level.
 Virtual desktop numbers start with 0.

Parameters (leading / can be omitted or - can be used instead):

/Help /h /?      this help screen.
/Verbose /Quiet  enable verbose (default) or quiet mode (short: /v and /q).
/Break /Continue break (default) or continue on error (short: /b and /co).
/List            list all virtual desktops (short: /li).
/Count           get count of virtual desktops to pipeline (short: /c).
/GetDesktop:<n|s> get number of virtual desktop <n> or desktop with text <s> in name to pipeline (short: /gd).
/GetCurrentDesktop  get number of current desktop to pipeline (short: /gcd).
/Name[:<s>] set name of desktop with number in pipeline (short: /na).
/Wallpaper[:<s>] set wallpaper path of desktop with number in pipeline (short: /wp)(only VirtualDesktop11.exe).
/AllWallpapers:<s> set wallpaper path of all desktops (short: /awp)(only VirtualDesktop11.exe).
/IsVisible[:<n|s>] is desktop number <n>, desktop with text <s> in name or with number in pipeline visible (short: /iv)? Returns 0 for visible and 1 for invisible.
/Switch[:<n|s>]  switch to desktop with number <n>, desktop with text <s> in name or with number in pipeline (short: /s).
/Left            switch to virtual desktop to the left of the active desktop (short: /l).
/Right           switch to virtual desktop to the right of the active desktop (short: /ri).
/Wrap /NoWrap /Left or /Right switch over or generate an error when the edge is reached (default)(short /w and /nw).
/New             create new desktop (short: /n). Number is stored in pipeline.
/Remove[:<n|s>]  remove desktop number <n>, desktop with text <s> in name or desktop with number in pipeline (short: /r).
/RemoveAll       remove all desktops but visible (short: /ra).
/SwapDesktop:<n|s>  swap desktop in pipeline with desktop number <n>, desktop with text <s> in name or desktop with number in pipeline (short: /sd).
/InsertDesktop:<n|s>  insert desktop number <n> or desktop with text <s> in name before desktop in pipeline or vice versa (short: /id)(not VirtualDesktop11.exe).
/MoveDesktop:<n|s>  move desktop in pipeline to desktop number <n> or desktop with text <s> in name (short: /md)(only VirtualDesktop11.exe).
/MoveWindowsToDesktop:<n|s>  move windows on desktop in pipeline to desktop number <n> or desktop with text <s> in name (short: /mwtd).
/MoveWindow:<s|n>  move process with name <s> or id <n> to desktop with number in pipeline (short: /mw).
/MoveWindowHandle:<s|n>  move window with text <s> in title or handle <n> to desktop with number in pipeline (short: /mwh).
/MoveActiveWindow move active window to desktop with number in pipeline (short: /maw).
/GetDesktopFromWindow:<s|n>  get desktop number where process with name <s> or id <n> is displayed (short: /gdfw).
/GetDesktopFromWindowHandle:<s|n>  get desktop number where window with text <s> in title or handle <n> is displayed (short: /gdfwh).
/IsWindowOnDesktop:<s|n>  check if process with name <s> or id <n> is on desktop with number in pipeline (short: /iwod). Returns 0 for yes, 1 for no.
/IsWindowHandleOnDesktop:<s|n>  check if window with text <s> in title or handle <n> is on desktop with number in pipeline (short: /iwhod). Returns 0 for yes, 1 for no.
/ListWindowsOnDesktop[:<n|s>]  list handles of windows on desktop number <n>, desktop with text <s> in name or desktop with number in pipeline (short: /lwod).
/CloseWindowsOnDesktop[:<n|s>]  close windows on desktop number <n>, desktop with text <s> in name or desktop with number in pipeline (short: /cwod).
/PinWindow:<s|n>   pin process with name <s> or id <n> to all desktops (short: /pw).
/PinWindowHandle:<s|n>   pin window with text <s> in title or handle <n> to all desktops (short: /pwh).
/UnPinWindow:<s|n>  unpin process with name <s> or id <n> from all desktops (short: /upw).
/UnPinWindowHandle:<s|n>  unpin window with text <s> in title or handle <n> from all desktops (short: /upwh).
/IsWindowPinned:<s|n>  check if process with name <s> or id <n> is pinned to all desktops (short: /iwp). Returns 0 for yes, 1 for no.
/IsWindowHandlePinned:<s|n>  check if window with text <s> in title or handle <n> is pinned to all desktops (short: /iwhp). Returns 0 for yes, 1 for no.
/PinApplication:<s|n>  pin application with name <s> or id <n> to all desktops (short: /pa).
/UnPinApplication:<s|n>  unpin application with name <s> or id <n> from all desktops (short: /upa).
/IsApplicationPinned:<s|n>  check if application with name <s> or id <n> is pinned to all desktops (short: /iap). Returns 0 for yes, 1 for no.
/Calc:<n>      add <n> to result, negative values are allowed (short: /ca).
/WaitKey       wait for key press (short: /wk).
/Sleep:<n>     wait for <n> milliseconds (short: /sl).

Hint: Instead of a desktop name you can use LAST or *LAST* to select the last virtual desktop.
Hint: Insert ^^ somewhere in window title parameters to prevent finding the own window. ^ is removed before searching window titles.


Examples:

Virtualdesktop.exe /LIST

Virtualdesktop.exe "-Switch:Desktop 2"

Virtualdesktop.exe -New -Switch -GetCurrentDesktop

Virtualdesktop.exe Q N /MOVEACTIVEWINDOW /SWITCH

Virtualdesktop.exe sleep:200 gd:1 mw:notepad s

Virtualdesktop.exe /Count /continue /Remove /Remove /Count

Virtualdesktop.exe /Count /Calc:-1 /Switch

VirtualDesktop.exe -IsWindowPinned:cmd
if ERRORLEVEL 1 VirtualDesktop.exe PinWindow:cmd

Virtualdesktop.exe -GetDesktop:*last* "-MoveWindowHandle:note^^pad"

for /f "tokens=4 delims= " %i in ('VirtualDesktop.exe c') do @set DesktopCount=%i
echo Count of desktops is %DesktopCount%
if %DesktopCount% GTR 1 VirtualDesktop.exe REMOVE
