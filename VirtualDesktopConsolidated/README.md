# VirtualDesktopConsolidated

**VirtualDesktopConsolidated** is a command-line tool for managing virtual desktops on Windows 10 and Windows 11 (including various builds and Windows Server editions). It provides a unified interface to control, automate, and query the Windows virtual desktop system, supporting a wide range of desktop and window management operations.

## Features

- List, create, remove, and switch between virtual desktops
- Set desktop names and (where supported) wallpapers
- Move windows and applications between desktops
- Pin and unpin windows or applications to all desktops
- Query and manipulate the state of desktops and windows
- Supports Windows 10, Windows 11 (21H2, 22H2, 24H2), and Windows Server (2016, 2019, 2022)
- Detects Windows version automatically and adapts API usage

## How It Works

The tool detects the current Windows version at runtime and uses the appropriate COM APIs to interact with the virtual desktop subsystem. It exposes a set of commands that can be chained or used individually to perform desktop management tasks. The result of each command can be used as input for subsequent commands, making it suitable for scripting and automation.

### Example Commands

- `VirtualDesktopC.exe list`  
  Lists all virtual desktops.
- `VirtualDesktopC.exe switch 2`  
  Switches to desktop number 2.
- `VirtualDesktopC.exe create`  
  Creates a new virtual desktop.
- `VirtualDesktopC.exe remove 1`  
  Removes desktop number 1.
- `VirtualDesktopC.exe setname 0 "Work"`  
  Sets the name of desktop 0 to "Work".
- `VirtualDesktopC.exe moveactive 1`  
  Moves the currently active window to desktop 1.
- `VirtualDesktopC.exe pinactive`  
  Pins the currently active window to all desktops.
- `VirtualDesktopC.exe iswindowpinned <window-title-or-pid>`  
  Checks if a window is pinned.

### Help

Run the executable with no arguments or with `/h`, `-h`, or `--help` to see the full list of supported commands and options.

## Building

The project targets .NET Framework 4.7.2. To build:

1. Open `VirtualDesktopConsolidated.sln` in Visual Studio.
2. Build the solution.

The output executable will be located in the `bin/Debug` or `bin/Release` directory under `src`.

## License

This project is licensed under the MIT License. 

## This solution is a work in progress and may not cover all edge cases or Windows versions. Contributions are welcome!

