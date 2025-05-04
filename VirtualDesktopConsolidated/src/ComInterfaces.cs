// COM interface definitions for all supported Windows versions
// This file is auto-generated as part of the VirtualDesktop consolidation plan

using System;
using System.Runtime.InteropServices;

namespace VirtualDesktop.Consolidated
{
    // Windows 10 (1809-22H2) interfaces
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("FF72FFDD-BE7E-43FC-9C03-AD81681E88E4")]
    internal interface IVirtualDesktop_Win10
    {
        bool IsViewVisible(IApplicationView_Win10 view);
        Guid GetId();
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("F31574D6-B682-4CDC-BD56-1827860ABEC6")]
    internal interface IVirtualDesktopManagerInternal_Win10
    {
        int GetCount();
        void MoveViewToDesktop(IApplicationView_Win10 view, IVirtualDesktop_Win10 desktop);
        bool CanViewMoveDesktops(IApplicationView_Win10 view);
        IVirtualDesktop_Win10 GetCurrentDesktop();
        void GetDesktops(out IObjectArray_Win10 desktops);
        [PreserveSig]
        int GetAdjacentDesktop(IVirtualDesktop_Win10 from, int direction, out IVirtualDesktop_Win10 desktop);
        void SwitchDesktop(IVirtualDesktop_Win10 desktop);
        IVirtualDesktop_Win10 CreateDesktop();
        void RemoveDesktop(IVirtualDesktop_Win10 desktop, IVirtualDesktop_Win10 fallback);
        IVirtualDesktop_Win10 FindDesktop(ref Guid desktopid);
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("1841C6D7-4F9D-42C0-AF41-8747538F10E5")]
    internal interface IApplicationViewCollection_Win10
    {
        int GetViews(out IObjectArray_Win10 array);
        int GetViewsByZOrder(out IObjectArray_Win10 array);
        int GetViewsByAppUserModelId(string id, out IObjectArray_Win10 array);
        int GetViewForHwnd(IntPtr hwnd, out IApplicationView_Win10 view);
        int GetViewForApplication(object application, out IApplicationView_Win10 view);
        int GetViewForAppUserModelId(string id, out IApplicationView_Win10 view);
        int GetViewInFocus(out IntPtr view);
        int Unknown1(out IntPtr view);
        void RefreshCollection();
        int RegisterForApplicationViewChanges(object listener, out int cookie);
        int UnregisterForApplicationViewChanges(int cookie);
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("9AC0B5C8-1484-4C5B-9533-4134A0F97CEA")]
    internal interface IApplicationView_Win10
    {
        // Methods omitted for brevity (see original source)
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("A5CD92FF-29BE-454C-8D04-D82879FB3F1B")]
    internal interface IVirtualDesktopManager_Win10
    {
        bool IsWindowOnCurrentVirtualDesktop(IntPtr topLevelWindow);
        Guid GetWindowDesktopId(IntPtr topLevelWindow);
        void MoveWindowToDesktop(IntPtr topLevelWindow, ref Guid desktopId);
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("4CE81583-1E4C-4632-A621-07A53543148F")]
    internal interface IVirtualDesktopPinnedApps_Win10
    {
        bool IsAppIdPinned(string appId);
        void PinAppID(string appId);
        void UnpinAppID(string appId);
        bool IsViewPinned(IApplicationView_Win10 applicationView);
        void PinView(IApplicationView_Win10 applicationView);
        void UnpinView(IApplicationView_Win10 applicationView);
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("92CA9DCD-5622-4BBA-A805-5E9F541BD8C9")]
    internal interface IObjectArray_Win10
    {
        void GetCount(out int count);
        void GetAt(int index, ref Guid iid, [MarshalAs(UnmanagedType.Interface)]out object obj);
    }

    // Windows 11 (21H2/22H2) interfaces
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("3F07F4BE-B107-441A-AF0F-39D82529072C")]
    internal interface IVirtualDesktop_Win11
    {
        bool IsViewVisible(IApplicationView_Win11 view);
        Guid GetId();
        [return: MarshalAs(UnmanagedType.HString)]
        string GetName();
        [return: MarshalAs(UnmanagedType.HString)]
        string GetWallpaperPath();
        bool IsRemote();
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("53F5CA0B-158F-4124-900C-057158060B27")]
    internal interface IVirtualDesktopManagerInternal_Win11
    {
        int GetCount();
        void MoveViewToDesktop(IApplicationView_Win11 view, IVirtualDesktop_Win11 desktop);
        bool CanViewMoveDesktops(IApplicationView_Win11 view);
        IVirtualDesktop_Win11 GetCurrentDesktop();
        void GetDesktops(out IObjectArray_Win11 desktops);
        [PreserveSig]
        int GetAdjacentDesktop(IVirtualDesktop_Win11 from, int direction, out IVirtualDesktop_Win11 desktop);
        void SwitchDesktop(IVirtualDesktop_Win11 desktop);
        IVirtualDesktop_Win11 CreateDesktop();
        void MoveDesktop(IVirtualDesktop_Win11 desktop, int nIndex);
        void RemoveDesktop(IVirtualDesktop_Win11 desktop, IVirtualDesktop_Win11 fallback);
        IVirtualDesktop_Win11 FindDesktop(ref Guid desktopid);
        void GetDesktopSwitchIncludeExcludeViews(IVirtualDesktop_Win11 desktop, out IObjectArray_Win11 unknown1, out IObjectArray_Win11 unknown2);
        void SetDesktopName(IVirtualDesktop_Win11 desktop, [MarshalAs(UnmanagedType.HString)] string name);
        void SetDesktopWallpaper(IVirtualDesktop_Win11 desktop, [MarshalAs(UnmanagedType.HString)] string path);
        void UpdateWallpaperPathForAllDesktops([MarshalAs(UnmanagedType.HString)] string path);
        void CopyDesktopState(IApplicationView_Win11 pView0, IApplicationView_Win11 pView1);
        void CreateRemoteDesktop([MarshalAs(UnmanagedType.HString)] string path, out IVirtualDesktop_Win11 desktop);
        void SwitchRemoteDesktop(IVirtualDesktop_Win11 desktop, IntPtr switchtype);
        void SwitchDesktopWithAnimation(IVirtualDesktop_Win11 desktop);
        void GetLastActiveDesktop(out IVirtualDesktop_Win11 desktop);
        void WaitForAnimationToComplete();
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("1841C6D7-4F9D-42C0-AF41-8747538F10E5")]
    internal interface IApplicationViewCollection_Win11
    {
        int GetViews(out IObjectArray_Win11 array);
        int GetViewsByZOrder(out IObjectArray_Win11 array);
        int GetViewsByAppUserModelId(string id, out IObjectArray_Win11 array);
        int GetViewForHwnd(IntPtr hwnd, out IApplicationView_Win11 view);
        int GetViewForApplication(object application, out IApplicationView_Win11 view);
        int GetViewForAppUserModelId(string id, out IApplicationView_Win11 view);
        int GetViewInFocus(out IntPtr view);
        int Unknown1(out IntPtr view);
        void RefreshCollection();
        int RegisterForApplicationViewChanges(object listener, out int cookie);
        int UnregisterForApplicationViewChanges(int cookie);
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("372E1D3B-38D3-42E4-A15B-8AB2B178F513")]
    internal interface IApplicationView_Win11
    {
        // Methods omitted for brevity (see original source)
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("A5CD92FF-29BE-454C-8D04-D82879FB3F1B")]
    internal interface IVirtualDesktopManager_Win11
    {
        bool IsWindowOnCurrentVirtualDesktop(IntPtr topLevelWindow);
        Guid GetWindowDesktopId(IntPtr topLevelWindow);
        void MoveWindowToDesktop(IntPtr topLevelWindow, ref Guid desktopId);
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("4CE81583-1E4C-4632-A621-07A53543148F")]
    internal interface IVirtualDesktopPinnedApps_Win11
    {
        bool IsAppIdPinned(string appId);
        void PinAppID(string appId);
        void UnpinAppID(string appId);
        bool IsViewPinned(IApplicationView_Win11 applicationView);
        void PinView(IApplicationView_Win11 applicationView);
        void UnpinView(IApplicationView_Win11 applicationView);
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("92CA9DCD-5622-4BBA-A805-5E9F541BD8C9")]
    internal interface IObjectArray_Win11
    {
        void GetCount(out int count);
        void GetAt(int index, ref Guid iid, [MarshalAs(UnmanagedType.Interface)]out object obj);
    }

    // Windows 11 24H2 interfaces
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("3F07F4BE-B107-441A-AF0F-39D82529072C")]
    internal interface IVirtualDesktop_Win11_24H2
    {
        bool IsViewVisible(IApplicationView_Win11_24H2 view);
        Guid GetId();
        [return: MarshalAs(UnmanagedType.HString)]
        string GetName();
        [return: MarshalAs(UnmanagedType.HString)]
        string GetWallpaperPath();
        bool IsRemote();
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("53F5CA0B-158F-4124-900C-057158060B27")]
    internal interface IVirtualDesktopManagerInternal_Win11_24H2
    {
        int GetCount();
        void MoveViewToDesktop(IApplicationView_Win11_24H2 view, IVirtualDesktop_Win11_24H2 desktop);
        bool CanViewMoveDesktops(IApplicationView_Win11_24H2 view);
        IVirtualDesktop_Win11_24H2 GetCurrentDesktop();
        void GetDesktops(out IObjectArray_Win11_24H2 desktops);
        [PreserveSig]
        int GetAdjacentDesktop(IVirtualDesktop_Win11_24H2 from, int direction, out IVirtualDesktop_Win11_24H2 desktop);
        void SwitchDesktop(IVirtualDesktop_Win11_24H2 desktop);
        void SwitchDesktopAndMoveForegroundView(IVirtualDesktop_Win11_24H2 desktop);
        IVirtualDesktop_Win11_24H2 CreateDesktop();
        void MoveDesktop(IVirtualDesktop_Win11_24H2 desktop, int nIndex);
        void RemoveDesktop(IVirtualDesktop_Win11_24H2 desktop, IVirtualDesktop_Win11_24H2 fallback);
        IVirtualDesktop_Win11_24H2 FindDesktop(ref Guid desktopid);
        void GetDesktopSwitchIncludeExcludeViews(IVirtualDesktop_Win11_24H2 desktop, out IObjectArray_Win11_24H2 unknown1, out IObjectArray_Win11_24H2 unknown2);
        void SetDesktopName(IVirtualDesktop_Win11_24H2 desktop, [MarshalAs(UnmanagedType.HString)] string name);
        void SetDesktopWallpaper(IVirtualDesktop_Win11_24H2 desktop, [MarshalAs(UnmanagedType.HString)] string path);
        void UpdateWallpaperPathForAllDesktops([MarshalAs(UnmanagedType.HString)] string path);
        void CopyDesktopState(IApplicationView_Win11_24H2 pView0, IApplicationView_Win11_24H2 pView1);
        void CreateRemoteDesktop([MarshalAs(UnmanagedType.HString)] string path, out IVirtualDesktop_Win11_24H2 desktop);
        void SwitchRemoteDesktop(IVirtualDesktop_Win11_24H2 desktop, IntPtr switchtype);
        void SwitchDesktopWithAnimation(IVirtualDesktop_Win11_24H2 desktop);
        void GetLastActiveDesktop(out IVirtualDesktop_Win11_24H2 desktop);
        void WaitForAnimationToComplete();
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("1841C6D7-4F9D-42C0-AF41-8747538F10E5")]
    internal interface IApplicationViewCollection_Win11_24H2
    {
        int GetViews(out IObjectArray_Win11_24H2 array);
        int GetViewsByZOrder(out IObjectArray_Win11_24H2 array);
        int GetViewsByAppUserModelId(string id, out IObjectArray_Win11_24H2 array);
        int GetViewForHwnd(IntPtr hwnd, out IApplicationView_Win11_24H2 view);
        int GetViewForApplication(object application, out IApplicationView_Win11_24H2 view);
        int GetViewForAppUserModelId(string id, out IApplicationView_Win11_24H2 view);
        int GetViewInFocus(out IntPtr view);
        int Unknown1(out IntPtr view);
        void RefreshCollection();
        int RegisterForApplicationViewChanges(object listener, out int cookie);
        int UnregisterForApplicationViewChanges(int cookie);
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("372E1D3B-38D3-42E4-A15B-8AB2B178F513")]
    internal interface IApplicationView_Win11_24H2
    {
        // Methods omitted for brevity (see original source)
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("A5CD92FF-29BE-454C-8D04-D82879FB3F1B")]
    internal interface IVirtualDesktopManager_Win11_24H2
    {
        bool IsWindowOnCurrentVirtualDesktop(IntPtr topLevelWindow);
        Guid GetWindowDesktopId(IntPtr topLevelWindow);
        void MoveWindowToDesktop(IntPtr topLevelWindow, ref Guid desktopId);
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("4CE81583-1E4C-4632-A621-07A53543148F")]
    internal interface IVirtualDesktopPinnedApps_Win11_24H2
    {
        bool IsAppIdPinned(string appId);
        void PinAppID(string appId);
        void UnpinAppID(string appId);
        bool IsViewPinned(IApplicationView_Win11_24H2 applicationView);
        void PinView(IApplicationView_Win11_24H2 applicationView);
        void UnpinView(IApplicationView_Win11_24H2 applicationView);
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("92CA9DCD-5622-4BBA-A805-5E9F541BD8C9")]
    internal interface IObjectArray_Win11_24H2
    {
        void GetCount(out int count);
        void GetAt(int index, ref Guid iid, [MarshalAs(UnmanagedType.Interface)]out object obj);
    }

    // Windows Server (2016/2019/2022) interfaces
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("FF72FFDD-BE7E-43FC-9C03-AD81681E88E4")]
    internal interface IVirtualDesktop_Server
    {
        bool IsViewVisible(IApplicationView_Server view);
        Guid GetId();
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("094afe11-44f2-4ba0-976f-29a97e263ee0")]
    internal interface IVirtualDesktopManagerInternal_Server
    {
        int GetCount(IntPtr hWndOrMon);
        void MoveViewToDesktop(IApplicationView_Server view, IVirtualDesktop_Server desktop);
        bool CanViewMoveDesktops(IApplicationView_Server view);
        IVirtualDesktop_Server GetCurrentDesktop(IntPtr hWndOrMon);
        void GetDesktops(IntPtr hWndOrMon, out IObjectArray_Server desktops);
        [PreserveSig]
        int GetAdjacentDesktop(IVirtualDesktop_Server from, int direction, out IVirtualDesktop_Server desktop);
        void SwitchDesktop(IntPtr hWndOrMon, IVirtualDesktop_Server desktop);
        IVirtualDesktop_Server CreateDesktop(IntPtr hWndOrMon);
        void RemoveDesktop(IVirtualDesktop_Server desktop, IVirtualDesktop_Server fallback);
        IVirtualDesktop_Server FindDesktop(ref Guid desktopid);
        void GetDesktopSwitchIncludeExcludeViews(IVirtualDesktop_Server desktop, out IObjectArray_Server unknown1, out IObjectArray_Server unknown2);
        void SetDesktopName(IVirtualDesktop_Server desktop, [MarshalAs(UnmanagedType.HString)] string name);
        void CopyDesktopState(IApplicationView_Server pView0, IApplicationView_Server pView1);
        int GetDesktopIsPerMonitor();
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("1841C6D7-4F9D-42C0-AF41-8747538F10E5")]
    internal interface IApplicationViewCollection_Server
    {
        int GetViews(out IObjectArray_Server array);
        int GetViewsByZOrder(out IObjectArray_Server array);
        int GetViewsByAppUserModelId(string id, out IObjectArray_Server array);
        int GetViewForHwnd(IntPtr hwnd, out IApplicationView_Server view);
        int GetViewForApplication(object application, out IApplicationView_Server view);
        int GetViewForAppUserModelId(string id, out IApplicationView_Server view);
        int GetViewInFocus(out IntPtr view);
        int Unknown1(out IntPtr view);
        void RefreshCollection();
        int RegisterForApplicationViewChanges(object listener, out int cookie);
        int UnregisterForApplicationViewChanges(int cookie);
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("9AC0B5C8-1484-4C5B-9533-4134A0F97CEA")]
    internal interface IApplicationView_Server
    {
        // Methods omitted for brevity (see original source)
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("A5CD92FF-29BE-454C-8D04-D82879FB3F1B")]
    internal interface IVirtualDesktopManager_Server
    {
        bool IsWindowOnCurrentVirtualDesktop(IntPtr topLevelWindow);
        Guid GetWindowDesktopId(IntPtr topLevelWindow);
        void MoveWindowToDesktop(IntPtr topLevelWindow, ref Guid desktopId);
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("4CE81583-1E4C-4632-A621-07A53543148F")]
    internal interface IVirtualDesktopPinnedApps_Server
    {
        bool IsAppIdPinned(string appId);
        void PinAppID(string appId);
        void UnpinAppID(string appId);
        bool IsViewPinned(IApplicationView_Server applicationView);
        void PinView(IApplicationView_Server applicationView);
        void UnpinView(IApplicationView_Server applicationView);
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("92CA9DCD-5622-4BBA-A805-5E9F541BD8C9")]
    internal interface IObjectArray_Server
    {
        void GetCount(out int count);
        void GetAt(int index, ref Guid iid, [MarshalAs(UnmanagedType.Interface)]out object obj);
    }

    // Placeholders for all COM interfaces, to be filled in with version-specific interfaces
    // Example:
    // [ComImport]
    // [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    // [Guid("...guid...")]
    // internal interface IVirtualDesktop_Win10 { ... }
    // [ComImport]
    // [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    // [Guid("...guid...")]
    // internal interface IVirtualDesktop_Win11_22H2 { ... }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("6D5140C1-7436-11CE-8034-00AA006009FA")]
    internal interface IServiceProvider
    {
        [PreserveSig]
        int QueryService(ref Guid guidService, ref Guid riid, out IntPtr ppvObject);
    }
} 