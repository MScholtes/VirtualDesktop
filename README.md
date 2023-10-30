# VirtualDesktop
C# command line tool to manage virtual desktops in Windows 10 and Windows 11

**a fork of [VirtualDesktop by MScholtes](https://github.com/MScholtes/VirtualDesktop);**

 Changes introduced

  - aliases `/NEXT` to `/RIGHT` and `/PREVIOUS` to `/LEFT`

  - adds new command `/JSON`
     * reports desktops in JSON parseable format

        ```{"count":1,"desktops":[{"name":"Desktop 1""visible":true,"wallpaper":"C:\\wallpaper.jpg"}]}```            
     * can be used by from the command line, but is mainly intended to be used in interactive mode

  
  - Interactive Mode  `/INTERACTIVE` or `/INT`
    
    * accepts keyboard input (or stdin via pipe, one command per line)
    * can be used for testing, but mainly implemented to allow [node.js modules](https://github.com/jonathan-annett/virtual-desktop-node) to use it
    * any changes that occur via keyboard/or other desktop managers are reported in JSON format:
       
       ```{"visibleIndex":2,"visible":"Desktop 3"}```
    * accepts all commands that can be used from the command line (ignores `/BREAK` and /`CONTINUE`, and `/INTERACTIVE`)

    * implements `/NAMES`, which is a JSON version of `/LIST` (only supported in interactive mode)

    * turns off `/VERBOSE` mode

    * `/NEW` will switch to the newly created desktop automatically

    * gives JSON responses to `/LEFT`, `/RIGHT`, `/NEXT`, `/PREVIOUS`, `/GCD`, `/NAMES`, `/NEW` (basically any command the in)

    * valid JSON responses will always be on their own line (preceded by and followed by \n)

    * doesn't filter out non-JSON responses, so anything reading piped output needs to keep this in mind.


original readme follows:

**Pre-compiled binaries in Releases now**

**Version 1.16, 2023-09-17**
- version for Windows 11 Insider Canary (build 25314 and up) called VirtualDesktop11InsiderCanary.cs
- (re)introduced parameter /RemoveAll for all versions

(look for a powershell version here: https://github.com/MScholtes/PSVirtualDesktop or here: https://www.powershellgallery.com/packages/VirtualDesktop)

**With Windows 11 22H2 Release 2215 Microsoft did change the API (COM GUIDs) for accessing the functions for virtual desktops again. I provide seven versions of virtualdesktop.cs now: virtualdesktop11.cs is for Windows 11 22H2 up to release 2134, virtualdesktop11-23h2.cs for Windows 11 22H2 release 2215 and newer (including Insider except Canary) versions, virtualdesktop11-21h2.cs for Windows 11 21H2, virtualdesktop11insidercanary.cs for Windows 11 Insider Canary, virtualdesktopserver2022.cs is for Windows Server 2022, virtualdesktop.cs is for Windows 10 1809 to 22H2, virtualdesktopserver2016.cs is for Windows Server 2016. Using Compile.bat all executables will be generated.**

**I will make a cleanup of versions with the next release!**

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

**/RemoveAll**       remove all desktops but visible (short: /ra).

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
