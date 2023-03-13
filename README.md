# VirtualDesktop
C# command line tool to manage virtual desktops in Windows 10 and Windows 11

**With Windows 11 Insider version again thanks to [NyaMisti](https://github.com/NyaMisty)!**

**Version 1.11, 2022-11-13**
**Version 1.12, 2023-03-10**
- new parameter /ListWindowsOnDesktop - list handles of windows on desktop
- new parameter /MoveWindowsToDesktop - move windows on desktop to another desktop
- new parameter /CloseWindowsOnDesktop -  close windows on desktop
- instead of a desktop name you can use LAST or \*LAST\* to select the last virtual desktop
- renamed version for Windows 10 1607 to Windows Server 2016
- removed version for Windows 10 1803

(look for a powershell version here: https://github.com/MScholtes/PSVirtualDesktop or here: https://www.powershellgallery.com/packages/VirtualDesktop)

**With Windows 11 22H2 Microsoft did change the API (COM GUIDs) for accessing the functions for virtual desktops again. I provide six versions of virtualdesktop.cs now: virtualdesktop11.cs is for Windows 11, virtualdesktop11-21h2.cs for Windows 11 21H2, virtualdesktopserver2022.cs is for Windows Server 2022, virtualdesktop.cs is for Windows 10 1809 to 21H2, virtualdesktopserver2016.cs is for Windows Server 2016. Using Compile.bat all executables will be generated.**

## Generate:
Compile with Compile.bat (no visual studio needed, but obviously Windows 10 or 11)

## Description:
Command line tool to manage the virtual desktops of Windows 10 and 11.
Parameters can be given as a sequence of commands. The result - most of the times the number of the processed desktop - can be used as input for the next parameter. The result of the last command is returned as error level.
Virtual desktop numbers start with 0.

## Parameters (leading / can be omitted or - can be used instead):
**/Help /h /?**      this help screen.

**/Verbose /Quiet**  enable verbose (default) or quiet mode (short: /v and /q).

**/Break /Continue** break (default) or continue on error (short: /b and /co).

**/List**            list all virtual desktops (short: /li).

**/Count**           get count of virtual desktops to pipeline (short: /c).

**/GetDesktop:&lt;n|s&gt;**  get number of virtual desktop &lt;n&gt; or desktop with text &lt;s&gt; in name to pipeline (short: /gd).

**/GetCurrentDesktop**  get number of current desktop to pipeline (short: /gcd).

**/Name[:&lt;s&gt;]**      set name of desktop with number in pipeline (short: /na).

**/Wallpaper[:&lt;s&gt;]**  set wallpaper path of desktop with number in pipeline (short: /wp)(only VirtualDesktop11.exe).

**/AllWallpapers:&lt;s&gt;**  set wallpaper path of all desktops (short: /awp)(only VirtualDesktop11.exe).

**/IsVisible[:&lt;n|s&gt;]**  is desktop number &lt;n&gt;, desktop with text &lt;s&gt; in name or number in pipeline visible (short: /iv)? Returns 0 for visible and 1 for invisible.

**/Switch[:&lt;n|s&gt;]**    switch to desktop with number &lt;n&gt;, desktop with text &lt;s&gt; in name or with number in pipeline (short: /s).

**/Left**            switch to virtual desktop to the left of the active desktop (short: /l).

**/Right**           switch to virtual desktop to the right of the active desktop (short: /ri).

**/Wrap /NoWrap**    /Left or /Right switch over or generate an error when the edge is reached (default)(short /w and /nw).

**/New**             create new desktop (short: /n). Number is stored in pipeline.

**/Remove[:&lt;n|s&gt;]**    remove desktop number &lt;n&gt;, desktop with text &lt;s&gt; in name or desktop with number in pipeline (short: /r).

**/RemoveAll**       remove all desktops but visible (short: /ra)(only VirtualDesktop11.exe).

**/SwapDesktop:&lt;n|s&gt;**  swap desktop in pipeline with desktop number &lt;n&gt;, desktop with text &lt;s&gt; in name or desktop with number in pipeline (short: /sd).

**/InsertDesktop:&lt;n|s&gt;**  insert desktop number &lt;n&gt; or desktop with text &lt;s&gt; in name before desktop in pipeline or vice versa (short: /id)(not VirtualDesktop11.exe).

**/MoveDesktop:&lt;n|s&gt;**  move desktop in pipeline to desktop number &lt;n&gt; or desktop with text &lt;s&gt; in name (short: /md)(only VirtualDesktop11.exe).

**/MoveWindowsToDesktop::&lt;n|s&gt;**  move windows on desktop in pipeline to desktop number &lt;n&gt; or desktop with text &lt;s&gt; in name (short: /mwtd).

**/MoveWindow:&lt;s|n&gt;**  move process with name &lt;s&gt; or id &lt;n&gt; to desktop with number in pipeline (short: /mw).

**/MoveWindowHandle:&lt;s|n&gt;**  move window with text &lt;s&gt; in title or handle &lt;n&gt; to desktop with number in pipeline (short: /mwh).

**/MoveActiveWindow**  move active window to desktop with number in pipeline (short: /maw).

**/GetDesktopFromWindow:&lt;s|n&gt;**  get desktop number where process with name &lt;s&gt; or id &lt;n&gt; is displayed (short: /gdfw).

**/GetDesktopFromWindowHandle:&lt;s|n&gt;**  get desktop number where window with text &lt;s&gt; in title or handle &lt;n&gt; is displayed (short: /gdfwh).

**/IsWindowOnDesktop:&lt;s|n&gt;**  check if process with name &lt;s&gt; or id &lt;n&gt; is on desktop with number in pipeline (short: /iwod). Returns 0 for yes, 1 for no.

**/IsWindowHandleOnDesktop:&lt;s|n&gt;**  check if window with text &lt;s&gt; in title or handle &lt;n&gt; is on desktop with number in pipeline (short: /iwhod). Returns 0 for yes, 1 for no.

**/ListWindowsOnDesktop[:&lt;n|s&gt;]**  list handles of windows on desktop number &lt;n&gt;, desktop with text &lt;s&gt; in name or desktop with number in pipeline (short: /lwod).

**/CloseWindowsOnDesktop[:&lt;n|s&gt;]**  close windows on desktop number &lt;n&gt;, desktop with text &lt;n&gt; in name or desktop with number in pipeline (short: /cwod).

**/PinWindow:&lt;s|n&gt;**   pin process with name &lt;s&gt; or id &lt;n&gt; to all desktops (short: /pw).

**/PinWindowHandle:&lt;s|n&gt;**   pin window with text &lt;s&gt; in title or handle &lt;n&gt; to all desktops (short: /pwh).

**/UnPinWindow:&lt;s|n&gt;**  unpin process with name &lt;s&gt; or id &lt;n&gt; from all desktops (short: /upw).

**/UnPinWindowHandle:&lt;s|n&gt;**  unpin window with text &lt;s&gt; in title or handle &lt;n&gt; from all desktops (short: /upwh).

**/IsWindowPinned:&lt;s|n&gt;**  check if process with name &lt;s&gt; or id &lt;n&gt; is pinned to all desktops (short: /iwp). Returns 0 for yes, 1 for no.

**/IsWindowHandlePinned:&lt;s|n&gt;**  check if window with text &lt;s&gt; in title or handle &lt;n&gt; is pinned to all desktops (short: /iwhp). Returns 0 for yes, 1 for no.

**/PinApplication:&lt;s|n&gt;**  pin application with name &lt;s&gt; or id &lt;n&gt; to all desktops (short: /pa).

**/UnPinApplication:&lt;s|n&gt;**  unpin application with name &lt;s&gt; or id &lt;n&gt; from all desktops (short: /upa).

**/IsApplicationPinned:&lt;s|n&gt;**  check if application with name &lt;s&gt; or id &lt;n&gt; is pinned to all desktops (short: /iap). Returns 0 for yes, 1 for no.

**/Calc:&lt;n&gt;**        add &lt;n&gt; to result, negative values are allowed (short: /ca).

**/WaitKey**       wait for key press (short: /wk).

**/Sleep:&lt;n&gt;**     wait for &lt;n&gt; milliseconds (short: /sl).

## Hints:
Instead of a desktop name you can use LAST or \*LAST\* to select the last virtual desktop.

Insert ^^ somewhere in window title parameters to prevent finding the own window. ^ is removed before searching window titles.

## Examples:
```bat
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
```
