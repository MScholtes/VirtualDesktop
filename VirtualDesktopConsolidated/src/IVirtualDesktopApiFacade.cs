using System;

namespace VirtualDesktop.Consolidated
{
    public interface IVirtualDesktopApiFacade
    {
        // Example methods (to be expanded with all needed methods)
        int GetDesktopCount();
        int GetCurrentDesktopIndex();
        void SwitchDesktop(int index);
        void CreateDesktop();
        void RemoveDesktop(int index, int fallbackIndex);
        void RemoveAllDesktopsExceptCurrent();
        string GetDesktopName(int index);
        void SetDesktopName(int index, string name);
        bool SupportsWallpaperSetting { get; }
        void SetDesktopWallpaper(int index, string path);
        void MoveWindowToDesktop(IntPtr hWnd, int desktopIndex);
        void MoveActiveWindowToDesktop(int desktopIndex);
        bool IsWindowOnDesktop(IntPtr hWnd, int desktopIndex);
        bool IsWindowPinned(IntPtr hWnd);
        void PinWindow(IntPtr hWnd);
        void PinActiveWindow();
        void UnpinWindow(IntPtr hWnd);
        void UnpinActiveWindow();
        bool IsApplicationPinned(IntPtr hWnd);
        void PinApplication(IntPtr hWnd);
        void UnpinApplication(IntPtr hWnd);
        int GetLeftDesktopIndex(int index);
        int GetRightDesktopIndex(int index);
        // Add other methods as needed for full feature parity
    }
} 