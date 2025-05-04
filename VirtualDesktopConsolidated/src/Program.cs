using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;

namespace VirtualDesktop.Consolidated
{
    static class Program
    {
        static int Main(string[] args)
        {
            Console.WriteLine("Detected Windows API Version: " + WindowsVersion.ApiVersion);
            if (WindowsVersion.ApiVersion == WindowsVersion.WindowsApiVersion.Unknown)
            {
                Console.WriteLine("Unsupported or undetected Windows version. Exiting.");
                return 1;
            }

            if (args.Length == 0 || args[0] == "/?" || args[0] == "-h" || args[0] == "--help")
            {
                HelpScreen();
                return 0;
            }

            try
            {
                switch (args[0].ToLower())
                {
                    case "list":
                        for (int i = 0; i < Desktop.Count; i++)
                        {
                            var d = new Desktop(i);
                            Console.WriteLine($"{i}: {d.Name}");
                        }
                        return 0;
                    case "switch":
                        if (args.Length > 1 && int.TryParse(args[1], out int switchIndex))
                        {
                            if (switchIndex < 0 || switchIndex >= Desktop.Count)
                                throw new ArgumentException($"Invalid desktop index: {switchIndex}");
                            new Desktop(switchIndex).MakeVisible();
                            return 0;
                        }
                        break;
                    case "create":
                        DesktopManager.ApiFacade.CreateDesktop();
                        return 0;
                    case "remove":
                        if (args.Length > 1 && int.TryParse(args[1], out int removeIndex))
                        {
                            int fallback = 0;
                            if (args.Length > 2 && int.TryParse(args[2], out int fb)) fallback = fb;
                            new Desktop(removeIndex).Remove(new Desktop(fallback));
                            return 0;
                        }
                        break;
                    case "removeallbutcurrent":
                        DesktopManager.ApiFacade.RemoveAllDesktopsExceptCurrent();
                        return 0;
                    case "setname":
                        if (args.Length > 2 && int.TryParse(args[1], out int nameIndex))
                        {
                            new Desktop(nameIndex).Name = string.Join(" ", args.Skip(2));
                            return 0;
                        }
                        break;
                    case "setwallpaper":
                        if (DesktopManager.ApiFacade.SupportsWallpaperSetting && args.Length > 2 && int.TryParse(args[1], out int wpIndex))
                        {
                            new Desktop(wpIndex).SetWallpaper(args[2]);
                            return 0;
                        }
                        else if (!DesktopManager.ApiFacade.SupportsWallpaperSetting)
                        {
                            Console.WriteLine("Wallpaper setting not supported on this Windows version.");
                            return 1;
                        }
                        break;
                    case "moveactive":
                        if (args.Length > 1 && int.TryParse(args[1], out int moveIndex))
                        {
                            DesktopManager.ApiFacade.MoveActiveWindowToDesktop(moveIndex);
                            return 0;
                        }
                        break;
                    case "pinactive":
                        DesktopManager.ApiFacade.PinActiveWindow();
                        return 0;
                    case "unpinactive":
                        DesktopManager.ApiFacade.UnpinActiveWindow();
                        return 0;
                    case "pinapp":
                        if (args.Length > 1)
                        {
                            var hWnd = ResolveWindowHandle(args[1]);
                            DesktopManager.ApiFacade.PinApplication(hWnd);
                            return 0;
                        }
                        break;
                    case "unpinapp":
                        if (args.Length > 1)
                        {
                            var hWnd = ResolveWindowHandle(args[1]);
                            DesktopManager.ApiFacade.UnpinApplication(hWnd);
                            return 0;
                        }
                        break;
                    case "iswindowpinned":
                        if (args.Length > 1)
                        {
                            var hWnd = ResolveWindowHandle(args[1]);
                            Console.WriteLine(DesktopManager.ApiFacade.IsWindowPinned(hWnd) ? "Pinned" : "Not pinned");
                            return 0;
                        }
                        break;
                    case "isapppinned":
                        if (args.Length > 1)
                        {
                            var hWnd = ResolveWindowHandle(args[1]);
                            Console.WriteLine(DesktopManager.ApiFacade.IsApplicationPinned(hWnd) ? "Pinned" : "Not pinned");
                            return 0;
                        }
                        break;
                    case "getleft":
                        if (args.Length > 1 && int.TryParse(args[1], out int leftIndex))
                        {
                            Console.WriteLine(DesktopManager.ApiFacade.GetLeftDesktopIndex(leftIndex));
                            return 0;
                        }
                        break;
                    case "getright":
                        if (args.Length > 1 && int.TryParse(args[1], out int rightIndex))
                        {
                            Console.WriteLine(DesktopManager.ApiFacade.GetRightDesktopIndex(rightIndex));
                            return 0;
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return 1;
            }

            Console.WriteLine("Unknown command or invalid arguments.");
            HelpScreen();
            return 1;
        }

        static void HelpScreen()
        {
            var version = WindowsVersion.ApiVersion;
            Console.WriteLine($"VirtualDesktop.exe\t\t\t\tMarkus Scholtes, Consolidated\n");
            Console.WriteLine("Command line tool to manage the virtual desktops of Windows 10 and 11.");
            Console.WriteLine("Parameters can be given as a sequence of commands. The result - most of the");
            Console.WriteLine("times the number of the processed desktop - can be used as input for the next");
            Console.WriteLine("parameter. The result of the last command is returned as error level.");
            Console.WriteLine("Virtual desktop numbers start with 0.\n");
            Console.WriteLine("Parameters (leading / can be omitted or - can be used instead):\n");
            Console.WriteLine("/Help /h /?      this help screen.");
            Console.WriteLine("/Verbose /Quiet  enable verbose (default) or quiet mode (short: /v and /q).");
            Console.WriteLine("/Break /Continue break (default) or continue on error (short: /b and /co).");
            if (version == VirtualDesktop.Consolidated.WindowsVersion.WindowsApiVersion.Windows11_21H2 ||
                version == VirtualDesktop.Consolidated.WindowsVersion.WindowsApiVersion.Windows11_22H2 ||
                version == VirtualDesktop.Consolidated.WindowsVersion.WindowsApiVersion.Windows11_24H2)
            {
                Console.WriteLine("/Animation:<s>   Enable switch animations (default) with 'On' or '1' and");
                Console.WriteLine("                   disable switch animations with 'Off' or '0' (short: /anim).");
            }
            Console.WriteLine("/List            list all virtual desktops (short: /li).");
            Console.WriteLine("/Count           get count of virtual desktops to pipeline (short: /c).");
            Console.WriteLine("/GetDesktop:<n|s> get number of virtual desktop <n> or desktop with text <s> in");
            Console.WriteLine("                   name to pipeline (short: /gd).");
            Console.WriteLine("/GetCurrentDesktop  get number of current desktop to pipeline (short: /gcd).");
            Console.WriteLine("/Name[:<s>]      set name of desktop with number in pipeline (short: /na).");
            if (DesktopManager.ApiFacade.SupportsWallpaperSetting)
            {
                Console.WriteLine("/Wallpaper[:<s>] set wallpaper path of desktop with number in pipeline (short:");
                Console.WriteLine("                   /wp).");
                Console.WriteLine("/AllWallpapers:<s> set wallpaper path of all desktops (short: /awp).");
            }
            Console.WriteLine("/IsVisible[:<n|s>] is desktop number <n>, desktop with text <s> in name or with");
            Console.WriteLine("                   number in pipeline visible (short: /iv)? Returns 0 for");
            Console.WriteLine("                   visible and 1 for invisible.");
            Console.WriteLine("/Switch[:<n|s>]  switch to desktop with number <n>, desktop with text <s> in");
            Console.WriteLine("                   name or with number in pipeline (short: /s).");
            Console.WriteLine("/Left            switch to virtual desktop to the left of the active desktop");
            Console.WriteLine("                   (short: /l).");
            Console.WriteLine("/Right           switch to virtual desktop to the right of the active desktop");
            Console.WriteLine("                   (short: /ri).");
            Console.WriteLine("/Wrap /NoWrap    /Left or /Right switch over or generate an error when the edge");
            Console.WriteLine("                   is reached (default)(short /w and /nw).");
            Console.WriteLine("/New             create new desktop (short: /n). Number is stored in pipeline.");
            Console.WriteLine("/Remove[:<n|s>]  remove desktop number <n>, desktop with text <s> in name or");
            Console.WriteLine("                   desktop with number in pipeline (short: /r).");
            Console.WriteLine("/RemoveAll       remove all desktops but visible (short: /ra).");
            Console.WriteLine("/SwapDesktop:<n|s>  swap desktop in pipeline with desktop number <n> or desktop");
            Console.WriteLine("                   with text <s> in name (short: /sd).");
            if (version != VirtualDesktop.Consolidated.WindowsVersion.WindowsApiVersion.Windows11_21H2 &&
                version != VirtualDesktop.Consolidated.WindowsVersion.WindowsApiVersion.Windows11_22H2 &&
                version != VirtualDesktop.Consolidated.WindowsVersion.WindowsApiVersion.Windows11_24H2)
            {
                Console.WriteLine("/InsertDesktop:<n|s>  insert desktop number <n> or desktop with text <s> in");
                Console.WriteLine("                   name before desktop in pipeline or vice versa (short: /id).");
            }
            if (version == VirtualDesktop.Consolidated.WindowsVersion.WindowsApiVersion.Windows11_21H2 ||
                version == VirtualDesktop.Consolidated.WindowsVersion.WindowsApiVersion.Windows11_22H2 ||
                version == VirtualDesktop.Consolidated.WindowsVersion.WindowsApiVersion.Windows11_24H2)
            {
                Console.WriteLine("/MoveDesktop:<n|s>  move desktop in pipeline to desktop number <n> or desktop");
                Console.WriteLine("                   with text <s> in name (short: /md).");
            }
            Console.WriteLine("/MoveWindowsToDesktop:<n|s>  move windows on desktop in pipeline to desktop");
            Console.WriteLine("                   number <n> or desktop with text <s> in name (short: /mwtd).");
            Console.WriteLine("/MoveWindow:<s|n>  move process with name <s> or id <n> to desktop with number");
            Console.WriteLine("                   in pipeline (short: /mw).");
            Console.WriteLine("/MoveWindowHandle:<s|n>  move window with text <s> in title or handle <n> to");
            Console.WriteLine("                   desktop with number in pipeline (short: /mwh).");
            Console.WriteLine("/MoveActiveWindow  move active window to desktop with number in pipeline");
            Console.WriteLine("                   (short: /maw).");
            Console.WriteLine("/GetDesktopFromWindow:<s|n>  get desktop number where process with name <s> or");
            Console.WriteLine("                   id <n> is displayed (short: /gdfw).");
            Console.WriteLine("/GetDesktopFromWindowHandle:<s|n>  get desktop number where window with text");
            Console.WriteLine("                   <s> in title or handle <n> is displayed (short: /gdfwh).");
            Console.WriteLine("/IsWindowOnDesktop:<s|n>  check if process with name <s> or id <n> is on");
            Console.WriteLine("                   desktop with number in pipeline (short: /iwod). Returns 0");
            Console.WriteLine("                   for yes, 1 for no.");
            Console.WriteLine("/IsWindowHandleOnDesktop:<s|n>  check if window with text <s> in title or");
            Console.WriteLine("                   handle <n> is on desktop with number in pipeline");
            Console.WriteLine("                   (short: /iwhod). Returns 0 for yes, 1 for no.");
            Console.WriteLine("/ListWindowsOnDesktop[:<n|s>]  list handles of windows on desktop number <n>,");
            Console.WriteLine("                   desktop with text <s> in name or desktop with number in");
            Console.WriteLine("                   pipeline (short: /lwod).");
            Console.WriteLine("/CloseWindowsOnDesktop[:<n|s>]  close windows on desktop number <n>, desktop");
            Console.WriteLine("                   with text <s> in name or desktop with number in pipeline");
            Console.WriteLine("                   (short: /cwod).");
            Console.WriteLine("/PinWindow:<s|n>  pin process with name <s> or id <n> to all desktops");
            Console.WriteLine("                   (short: /pw).");
            Console.WriteLine("/PinActiveWindow  pin active window to all desktops (short: /paw).");
            Console.WriteLine("/PinWindowHandle:<s|n>  pin window with text <s> in title or handle <n> to all");
            Console.WriteLine("                   desktops (short: /pwh).");
            Console.WriteLine("/UnPinWindow:<s|n>  unpin process with name <s> or id <n> from all desktops");
            Console.WriteLine("                   (short: /upw).");
            Console.WriteLine("/UnPinActiveWindow  unpin active window from all desktops (short: /upaw).");
            Console.WriteLine("/UnPinWindowHandle:<s|n>  unpin window with text <s> in title or handle <n>");
            Console.WriteLine("                   from all desktops (short: /upwh).");
            Console.WriteLine("/IsWindowPinned:<s|n>  check if process with name <s> or id <n> is pinned to");
            Console.WriteLine("                   all desktops (short: /iwp). Returns 0 for yes, 1 for no.");
            Console.WriteLine("/IsWindowHandlePinned:<s|n>  check if window with text <s> in title or handle");
            Console.WriteLine("                   <n> is pinned to all desktops (short: /iwhp). Returns 0 for");
            Console.WriteLine("                   yes, 1 for no.");
            Console.WriteLine("/PinApplication:<s|n>  pin application with name <s> or id <n> to all desktops");
            Console.WriteLine("                   (short: /pa).");
            Console.WriteLine("/UnPinApplication:<s|n>  unpin application with name <s> or id <n> from all");
            Console.WriteLine("                   desktops (short: /upa).");
            Console.WriteLine("/IsApplicationPinned:<s|n>  check if application with name <s> or id <n> is");
            Console.WriteLine("                   pinned to all desktops (short: /iap). Returns 0 for yes, 1");
            Console.WriteLine("                   for no.");
            Console.WriteLine("/Calc:<n>        add <n> to result, negative values are allowed (short: /ca).");
            Console.WriteLine("/WaitKey         wait for key press (short: /wk).");
            Console.WriteLine("/Sleep:<n>       wait for <n> milliseconds (short: /sl).\n");
            Console.WriteLine("Hint: Instead of a desktop name you can use LAST or *LAST* to select the last");
            Console.WriteLine("virtual desktop.");
            Console.WriteLine("Hint: Insert ^^ somewhere in window title parameters to prevent finding the own");
            Console.WriteLine("window. ^ is removed before searching window titles.\n");
            Console.WriteLine("Examples:");
            Console.WriteLine("Virtualdesktop.exe /LIST");
            Console.WriteLine("Virtualdesktop.exe \"-Switch:Desktop 2\"");
            Console.WriteLine("Virtualdesktop.exe -New -Switch -GetCurrentDesktop");
            Console.WriteLine("Virtualdesktop.exe Q N /MOVEACTIVEWINDOW /SWITCH");
            Console.WriteLine("Virtualdesktop.exe sleep:200 gd:1 mw:notepad s");
            Console.WriteLine("Virtualdesktop.exe /Count /continue /Remove /Remove /Count");
            Console.WriteLine("Virtualdesktop.exe /Count /Calc:-1 /Switch");
            Console.WriteLine("VirtualDesktop.exe -IsWindowPinned:cmd");
            Console.WriteLine("if ERRORLEVEL 1 VirtualDesktop.exe PinWindow:cmd");
            Console.WriteLine("Virtualdesktop.exe -GetDesktop:*last* \"-MoveWindowHandle:note^^pad\"");
        }

        static IntPtr ResolveWindowHandle(string arg)
        {
            // Try PID
            if (int.TryParse(arg, out int pid))
            {
                var proc = Process.GetProcessById(pid);
                return proc.MainWindowHandle;
            }
            // Try window title
            foreach (var proc in Process.GetProcesses())
            {
                if (!string.IsNullOrEmpty(proc.MainWindowTitle) && proc.MainWindowTitle.IndexOf(arg, StringComparison.OrdinalIgnoreCase) >= 0)
                    return proc.MainWindowHandle;
            }
            throw new ArgumentException($"Could not resolve window handle for '{arg}'");
        }
    }
} 