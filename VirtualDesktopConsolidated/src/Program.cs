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

            // State variables
            bool verbose = true;
            bool breakonerror = true;
            bool wrapdesktops = false;
            int rc = 0;

            // Helper for error output
            void PrintError(string msg)
            {
                Console.Error.WriteLine(msg);
            }

            // Helper for verbose output
            void PrintVerbose(string msg)
            {
                if (verbose) Console.WriteLine(msg);
            }

            // Argument parsing and command loop
            foreach (string arg in args)
            {
                // Parse argument: support /, -, or no prefix, and : or = as separator
                var match = System.Text.RegularExpressions.Regex.Match(arg, @"^[-\/]?([^:=]+)[:=]?(.*)$");
                if (match.Groups.Count != 3)
                {
                    rc = -2;
                    PrintError($"Error in parameter '{arg}'");
                    if (breakonerror) break;
                    else continue;
                }
                string cmd = match.Groups[1].Value.Trim().ToUpper();
                string val = match.Groups[2].Value;

                // Reset rc if previously errored
                if (rc < 0) rc = 0;

                // Commands without value
                if (string.IsNullOrEmpty(val))
                {
                    switch (cmd)
                    {
                        case "HELP": case "H": case "?":
                            HelpScreen();
                            return 0;
                        case "QUIET": case "Q":
                            verbose = false;
                            break;
                        case "VERBOSE": case "V":
                            PrintVerbose("Verbose mode enabled");
                            verbose = true;
                            break;
                        case "BREAK": case "B":
                            if (verbose) Console.WriteLine("Break on error enabled");
                            breakonerror = true;
                            break;
                        case "CONTINUE": case "CO":
                            if (verbose) Console.WriteLine("Break on error disabled");
                            breakonerror = false;
                            break;
                        case "WRAP": case "W":
                            if (verbose) Console.WriteLine("Wrapping desktops enabled");
                            wrapdesktops = true;
                            break;
                        case "NOWRAP": case "NW":
                            if (verbose) Console.WriteLine("Wrapping desktop disabled");
                            wrapdesktops = false;
                            break;
                        case "COUNT": case "C":
                            rc = Desktop.Count;
                            PrintVerbose($"Count of desktops: {rc}");
                            break;
                        case "LIST": case "LI":
                            int desktopCount = Desktop.Count;
                            int visibleDesktop = Desktop.Current.Index;
                            if (verbose)
                            {
                                Console.WriteLine("Virtual desktops:");
                                Console.WriteLine("-----------------");
                            }
                            for (int i = 0; i < desktopCount; i++)
                            {
                                string name = DesktopManager.ApiFacade.GetDesktopName(i);
                                if (i != visibleDesktop)
                                    Console.Write(name);
                                else
                                    Console.Write(name + " (visible)");
                                if (DesktopManager.ApiFacade.SupportsWallpaperSetting)
                                {
                                    string wppath = Desktop.DesktopWallpaperFromIndex(i);
                                    if (!string.IsNullOrEmpty(wppath))
                                        Console.WriteLine($" (Wallpaper: {wppath})");
                                    else
                                        Console.WriteLine();
                                }
                                else
                                    Console.WriteLine();
                            }
                            if (verbose) Console.WriteLine($"\nCount of desktops: {desktopCount}");
                            break;
                        case "GETCURRENTDESKTOP": case "GCD":
                            rc = Desktop.Current.Index;
                            PrintVerbose($"Current desktop: '{DesktopManager.ApiFacade.GetDesktopName(rc)}' (desktop number {rc})");
                            break;
                        case "NEW": case "N":
                            try
                            {
                                DesktopManager.ApiFacade.CreateDesktop();
                                rc = Desktop.Count - 1;
                                PrintVerbose($"Created new desktop, number {rc}");
                            }
                            catch
                            {
                                rc = -1;
                                PrintError("Error creating new desktop");
                            }
                            break;
                        case "REMOVEALL": case "RA":
                            try
                            {
                                DesktopManager.ApiFacade.RemoveAllDesktopsExceptCurrent();
                                PrintVerbose("Removed all desktops but visible");
                            }
                            catch
                            {
                                rc = -1;
                                PrintError("Error removing all desktops but visible");
                            }
                            break;
                        case "LEFT": case "L":
                            try
                            {
                                int left = DesktopManager.ApiFacade.GetLeftDesktopIndex(Desktop.Current.Index);
                                if (wrapdesktops && left == -1)
                                    left = Desktop.Count - 1;
                                if (left >= 0)
                                {
                                    DesktopManager.ApiFacade.SwitchDesktop(left);
                                    rc = left;
                                    PrintVerbose($"Switched to left desktop {left}");
                                }
                                else
                                {
                                    rc = -1;
                                    PrintError("No left desktop");
                                }
                            }
                            catch
                            {
                                rc = -1;
                                PrintError("Error switching to left desktop");
                            }
                            break;
                        case "RIGHT": case "RI":
                            try
                            {
                                int right = DesktopManager.ApiFacade.GetRightDesktopIndex(Desktop.Current.Index);
                                if (wrapdesktops && right == -1)
                                    right = 0;
                                if (right >= 0)
                                {
                                    DesktopManager.ApiFacade.SwitchDesktop(right);
                                    rc = right;
                                    PrintVerbose($"Switched to right desktop {right}");
                                }
                                else
                                {
                                    rc = -1;
                                    PrintError("No right desktop");
                                }
                            }
                            catch
                            {
                                rc = -1;
                                PrintError("Error switching to right desktop");
                            }
                            break;
                        case "SWITCH": case "S":
                            int iParam = -1;
                            if (int.TryParse(val, out iParam))
                            {
                                try
                                {
                                    DesktopManager.ApiFacade.SwitchDesktop(iParam);
                                    rc = iParam;
                                    PrintVerbose($"Switched to desktop {iParam}");
                                }
                                catch
                                {
                                    rc = -1;
                                    PrintError("Error switching desktop");
                                }
                            }
                            else if (val.Trim().ToUpper() == "LAST" || val.Trim().ToUpper() == "*LAST*")
                            {
                                iParam = Desktop.Count - 1;
                                try
                                {
                                    DesktopManager.ApiFacade.SwitchDesktop(iParam);
                                    rc = iParam;
                                    PrintVerbose($"Switched to last desktop {iParam}");
                                }
                                catch
                                {
                                    rc = -1;
                                    PrintError("Error switching to last desktop");
                                }
                            }
                            else
                            {
                                iParam = -1;
                                for (int i = Desktop.Count - 1; i >= 0; i--)
                                {
                                    if (DesktopManager.ApiFacade.GetDesktopName(i).ToUpper().Contains(val.Trim().ToUpper()))
                                    {
                                        iParam = i;
                                        break;
                                    }
                                }
                                if (iParam >= 0)
                                {
                                    try
                                    {
                                        DesktopManager.ApiFacade.SwitchDesktop(iParam);
                                        rc = iParam;
                                        PrintVerbose($"Switched to desktop {iParam} ('{DesktopManager.ApiFacade.GetDesktopName(iParam)}')");
                                    }
                                    catch
                                    {
                                        rc = -1;
                                        PrintError("Error switching desktop by name");
                                    }
                                }
                                else
                                {
                                    rc = -2;
                                    PrintError($"Could not find virtual desktop with name containing '{val}'");
                                }
                            }
                            break;
                        case "REMOVE": case "R":
                            iParam = -1;
                            if (int.TryParse(val, out iParam))
                            {
                                try
                                {
                                    DesktopManager.ApiFacade.RemoveDesktop(iParam, 0);
                                    rc = iParam;
                                    PrintVerbose($"Removed desktop {iParam}");
                                }
                                catch
                                {
                                    rc = -1;
                                    PrintError("Error removing desktop");
                                }
                            }
                            else if (val.Trim().ToUpper() == "LAST" || val.Trim().ToUpper() == "*LAST*")
                            {
                                iParam = Desktop.Count - 1;
                                try
                                {
                                    DesktopManager.ApiFacade.RemoveDesktop(iParam, 0);
                                    rc = iParam;
                                    PrintVerbose($"Removed last desktop {iParam}");
                                }
                                catch
                                {
                                    rc = -1;
                                    PrintError("Error removing last desktop");
                                }
                            }
                            else
                            {
                                iParam = -1;
                                for (int i = Desktop.Count - 1; i >= 0; i--)
                                {
                                    if (DesktopManager.ApiFacade.GetDesktopName(i).ToUpper().Contains(val.Trim().ToUpper()))
                                    {
                                        iParam = i;
                                        break;
                                    }
                                }
                                if (iParam >= 0)
                                {
                                    try
                                    {
                                        DesktopManager.ApiFacade.RemoveDesktop(iParam, 0);
                                        rc = iParam;
                                        PrintVerbose($"Removed desktop {iParam} ('{DesktopManager.ApiFacade.GetDesktopName(iParam)}')");
                                    }
                                    catch
                                    {
                                        rc = -1;
                                        PrintError("Error removing desktop by name");
                                    }
                                }
                                else
                                {
                                    rc = -2;
                                    PrintError($"Could not find virtual desktop with name containing '{val}'");
                                }
                            }
                            break;
                        case "PINACTIVEWINDOW": case "PAW":
                            try
                            {
                                DesktopManager.ApiFacade.PinActiveWindow();
                                PrintVerbose("Active window pinned");
                            }
                            catch
                            {
                                rc = -1;
                                PrintError("Pinning of active window failed");
                            }
                            break;
                        case "UNPINACTIVEWINDOW": case "UPAW":
                            try
                            {
                                DesktopManager.ApiFacade.UnpinActiveWindow();
                                PrintVerbose("Active window unpinned");
                            }
                            catch
                            {
                                rc = -1;
                                PrintError("Unpinning of active window failed");
                            }
                            break;
                        // Add more no-value commands as needed
                        default:
                            rc = -2;
                            PrintError($"Error in parameter '{arg}'");
                            if (breakonerror) break;
                            continue;
                    }
                }
                else // Commands with value
                {
                    int iParam = -1;
                    switch (cmd)
                    {
                        case "ANIMATION": case "ANIM":
                            switch (val.Trim().ToUpper())
                            {
                                case "ON": case "1":
                                    // TODO: Implement animation enable
                                    PrintVerbose("Enabled switch animations");
                                    break;
                                case "OFF": case "0":
                                    // TODO: Implement animation disable
                                    PrintVerbose("Disabled switch animations");
                                    break;
                                default:
                                    rc = -2;
                                    break;
                            }
                            break;
                        case "GETDESKTOP": case "GD":
                            if (int.TryParse(val, out iParam))
                            {
                                if (iParam >= 0 && iParam < Desktop.Count)
                                {
                                    PrintVerbose($"Virtual desktop number {iParam} ('{DesktopManager.ApiFacade.GetDesktopName(iParam)}') selected");
                                    rc = iParam;
                                }
                                else rc = -1;
                            }
                            else
                            {
                                // Search by name or LAST
                                if (val.Trim().ToUpper() == "LAST" || val.Trim().ToUpper() == "*LAST*")
                                {
                                    iParam = Desktop.Count - 1;
                                    PrintVerbose($"Virtual desktop number {iParam} ('{DesktopManager.ApiFacade.GetDesktopName(iParam)}') selected");
                                    rc = iParam;
                                }
                                else
                                {
                                    // Search by partial name
                                    iParam = -1;
                                    for (int i = Desktop.Count - 1; i >= 0; i--)
                                    {
                                        if (DesktopManager.ApiFacade.GetDesktopName(i).ToUpper().Contains(val.Trim().ToUpper()))
                                        {
                                            iParam = i;
                                            break;
                                        }
                                    }
                                    if (iParam >= 0)
                                    {
                                        PrintVerbose($"Virtual desktop number {iParam} ('{DesktopManager.ApiFacade.GetDesktopName(iParam)}') selected");
                                        rc = iParam;
                                    }
                                    else
                                    {
                                        PrintVerbose($"Could not find virtual desktop with name containing '{val}'");
                                        rc = -2;
                                    }
                                }
                            }
                            break;
                        case "NAME": case "NA":
                            try
                            {
                                DesktopManager.ApiFacade.SetDesktopName(rc, val);
                                PrintVerbose($"Set name of desktop {rc} to '{val}'");
                            }
                            catch
                            {
                                rc = -1;
                                PrintError("Error setting desktop name");
                            }
                            break;
                        case "WALLPAPER": case "WP":
                            try
                            {
                                DesktopManager.ApiFacade.SetDesktopWallpaper(rc, val);
                                PrintVerbose($"Set wallpaper of desktop {rc} to '{val}'");
                            }
                            catch
                            {
                                rc = -1;
                                PrintError("Error setting wallpaper");
                            }
                            break;
                        case "ALLWALLPAPERS": case "AWP":
                            try
                            {
                                for (int i = 0; i < Desktop.Count; i++)
                                    DesktopManager.ApiFacade.SetDesktopWallpaper(i, val);
                                PrintVerbose($"Set wallpaper of all desktops to '{val}'");
                            }
                            catch
                            {
                                rc = -1;
                                PrintError("Error setting wallpaper for all desktops");
                            }
                            break;
                        case "SWITCH": case "S":
                            if (int.TryParse(val, out iParam))
                            {
                                try
                                {
                                    DesktopManager.ApiFacade.SwitchDesktop(iParam);
                                    rc = iParam;
                                    PrintVerbose($"Switched to desktop {iParam}");
                                }
                                catch
                                {
                                    rc = -1;
                                    PrintError("Error switching desktop");
                                }
                            }
                            else if (val.Trim().ToUpper() == "LAST" || val.Trim().ToUpper() == "*LAST*")
                            {
                                iParam = Desktop.Count - 1;
                                try
                                {
                                    DesktopManager.ApiFacade.SwitchDesktop(iParam);
                                    rc = iParam;
                                    PrintVerbose($"Switched to last desktop {iParam}");
                                }
                                catch
                                {
                                    rc = -1;
                                    PrintError("Error switching to last desktop");
                                }
                            }
                            else
                            {
                                iParam = -1;
                                for (int i = Desktop.Count - 1; i >= 0; i--)
                                {
                                    if (DesktopManager.ApiFacade.GetDesktopName(i).ToUpper().Contains(val.Trim().ToUpper()))
                                    {
                                        iParam = i;
                                        break;
                                    }
                                }
                                if (iParam >= 0)
                                {
                                    try
                                    {
                                        DesktopManager.ApiFacade.SwitchDesktop(iParam);
                                        rc = iParam;
                                        PrintVerbose($"Switched to desktop {iParam} ('{DesktopManager.ApiFacade.GetDesktopName(iParam)}')");
                                    }
                                    catch
                                    {
                                        rc = -1;
                                        PrintError("Error switching desktop by name");
                                    }
                                }
                                else
                                {
                                    rc = -2;
                                    PrintError($"Could not find virtual desktop with name containing '{val}'");
                                }
                            }
                            break;
                        case "REMOVE": case "R":
                            if (int.TryParse(val, out iParam))
                            {
                                try
                                {
                                    DesktopManager.ApiFacade.RemoveDesktop(iParam, 0);
                                    rc = iParam;
                                    PrintVerbose($"Removed desktop {iParam}");
                                }
                                catch
                                {
                                    rc = -1;
                                    PrintError("Error removing desktop");
                                }
                            }
                            else if (val.Trim().ToUpper() == "LAST" || val.Trim().ToUpper() == "*LAST*")
                            {
                                iParam = Desktop.Count - 1;
                                try
                                {
                                    DesktopManager.ApiFacade.RemoveDesktop(iParam, 0);
                                    rc = iParam;
                                    PrintVerbose($"Removed last desktop {iParam}");
                                }
                                catch
                                {
                                    rc = -1;
                                    PrintError("Error removing last desktop");
                                }
                            }
                            else
                            {
                                iParam = -1;
                                for (int i = Desktop.Count - 1; i >= 0; i--)
                                {
                                    if (DesktopManager.ApiFacade.GetDesktopName(i).ToUpper().Contains(val.Trim().ToUpper()))
                                    {
                                        iParam = i;
                                        break;
                                    }
                                }
                                if (iParam >= 0)
                                {
                                    try
                                    {
                                        DesktopManager.ApiFacade.RemoveDesktop(iParam, 0);
                                        rc = iParam;
                                        PrintVerbose($"Removed desktop {iParam} ('{DesktopManager.ApiFacade.GetDesktopName(iParam)}')");
                                    }
                                    catch
                                    {
                                        rc = -1;
                                        PrintError("Error removing desktop by name");
                                    }
                                }
                                else
                                {
                                    rc = -2;
                                    PrintError($"Could not find virtual desktop with name containing '{val}'");
                                }
                            }
                            break;
                        // Add more value commands as needed (MOVEWINDOW, PINWINDOW, etc.)
                        default:
                            rc = -2;
                            PrintError($"Error in parameter '{arg}'");
                            if (breakonerror) break;
                            continue;
                    }
                }

                if (rc == -1)
                {
                    PrintError($"Error while processing '{arg}'");
                    if (breakonerror) break;
                }
                if (rc == -2)
                {
                    PrintError($"Error in parameter '{arg}'");
                    if (breakonerror) break;
                }
            }

            return rc;
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
            Console.WriteLine("/IsWindowHandleOnDesktop:<s|n>  check if window with text <s> in title or handle");
            Console.WriteLine("                   <n> is on desktop with number in pipeline");
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