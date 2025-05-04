using System;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace VirtualDesktop.Consolidated
{
    public class VirtualDesktopApiFacade_Win11_24H2 : IVirtualDesktopApiFacade
    {
        private readonly IVirtualDesktopManagerInternal_Win11_24H2 _managerInternal;
        private readonly IVirtualDesktopManager_Win11_24H2 _manager;
        private readonly IApplicationViewCollection_Win11_24H2 _viewCollection;
        private readonly IVirtualDesktopPinnedApps_Win11_24H2 _pinnedApps;

        public VirtualDesktopApiFacade_Win11_24H2()
        {
            var immersiveShellType = Type.GetTypeFromCLSID(new Guid("C2F03A33-21F5-47FA-B4BB-156362A2F239"));
            var shell = (IServiceProvider)Activator.CreateInstance(immersiveShellType);
            Guid serviceGuid = new Guid("C5E0CDCA-7B6E-41B2-9FC4-D93975CC467B");
            Guid iid = typeof(IVirtualDesktopManagerInternal_Win11_24H2).GUID;
            IntPtr ppv;
            int hr = shell.QueryService(ref serviceGuid, ref iid, out ppv);
            if (hr != 0 || ppv == IntPtr.Zero)
                throw new InvalidCastException("QueryService did not return IVirtualDesktopManagerInternal_Win11_24H2");
            _managerInternal = (IVirtualDesktopManagerInternal_Win11_24H2)Marshal.GetObjectForIUnknown(ppv);
            _manager = (IVirtualDesktopManager_Win11_24H2)Activator.CreateInstance(Type.GetTypeFromCLSID(new Guid("AA509086-5CA9-4C25-8F95-589D3C07B48A")));
            iid = typeof(IApplicationViewCollection_Win11_24H2).GUID;
            hr = shell.QueryService(ref iid, ref iid, out ppv);
            if (hr != 0 || ppv == IntPtr.Zero)
                throw new InvalidCastException("QueryService did not return IApplicationViewCollection_Win11_24H2");
            _viewCollection = (IApplicationViewCollection_Win11_24H2)Marshal.GetObjectForIUnknown(ppv);
            serviceGuid = new Guid("B5A399E7-1C87-46B8-88E9-FC5747B171BD");
            iid = typeof(IVirtualDesktopPinnedApps_Win11_24H2).GUID;
            hr = shell.QueryService(ref serviceGuid, ref iid, out ppv);
            if (hr != 0 || ppv == IntPtr.Zero)
                throw new InvalidCastException("QueryService did not return IVirtualDesktopPinnedApps_Win11_24H2");
            _pinnedApps = (IVirtualDesktopPinnedApps_Win11_24H2)Marshal.GetObjectForIUnknown(ppv);
        }

        public int GetDesktopCount() => _managerInternal.GetCount();

        public int GetCurrentDesktopIndex()
        {
            var current = _managerInternal.GetCurrentDesktop();
            _managerInternal.GetDesktops(out var desktops);
            desktops.GetCount(out int count);
            for (int i = 0; i < count; i++)
            {
                desktops.GetAt(i, typeof(IVirtualDesktop_Win11_24H2).GUID, out var obj);
                if (((IVirtualDesktop_Win11_24H2)obj).GetId() == current.GetId())
                    return i;
            }
            return -1;
        }

        public void SwitchDesktop(int index)
        {
            _managerInternal.GetDesktops(out var desktops);
            desktops.GetCount(out int count);
            if (index < 0 || index >= count) throw new ArgumentOutOfRangeException(nameof(index));
            desktops.GetAt(index, typeof(IVirtualDesktop_Win11_24H2).GUID, out var obj);
            _managerInternal.SwitchDesktop((IVirtualDesktop_Win11_24H2)obj);
        }

        public void CreateDesktop() => _managerInternal.CreateDesktop();

        public void RemoveDesktop(int index, int fallbackIndex)
        {
            _managerInternal.GetDesktops(out var desktops);
            desktops.GetCount(out int count);
            if (index < 0 || index >= count) throw new ArgumentOutOfRangeException(nameof(index));
            if (fallbackIndex < 0 || fallbackIndex >= count) throw new ArgumentOutOfRangeException(nameof(fallbackIndex));
            desktops.GetAt(index, typeof(IVirtualDesktop_Win11_24H2).GUID, out var objRemove);
            desktops.GetAt(fallbackIndex, typeof(IVirtualDesktop_Win11_24H2).GUID, out var objFallback);
            _managerInternal.RemoveDesktop((IVirtualDesktop_Win11_24H2)objRemove, (IVirtualDesktop_Win11_24H2)objFallback);
        }

        public void RemoveAllDesktopsExceptCurrent()
        {
            int desktopCount = GetDesktopCount();
            int current = GetCurrentDesktopIndex();
            if (current < desktopCount - 1)
            {
                for (int i = desktopCount - 1; i > current; i--)
                    RemoveDesktop(i, current);
            }
            if (current > 0)
            {
                for (int i = 0; i < current; i++)
                    RemoveDesktop(0, current - i - 1);
            }
        }

        public string GetDesktopName(int index)
        {
            _managerInternal.GetDesktops(out var desktops);
            desktops.GetAt(index, typeof(IVirtualDesktop_Win11_24H2).GUID, out var objDesktop);
            try { return ((IVirtualDesktop_Win11_24H2)objDesktop).GetName(); } catch { return $"Desktop {index + 1}"; }
        }

        public void SetDesktopName(int index, string name)
        {
            _managerInternal.GetDesktops(out var desktops);
            desktops.GetAt(index, typeof(IVirtualDesktop_Win11_24H2).GUID, out var objDesktop);
            _managerInternal.SetDesktopName((IVirtualDesktop_Win11_24H2)objDesktop, name);
        }

        public bool SupportsWallpaperSetting => true;
        public void SetDesktopWallpaper(int index, string path)
        {
            _managerInternal.GetDesktops(out var desktops);
            desktops.GetAt(index, typeof(IVirtualDesktop_Win11_24H2).GUID, out var objDesktop);
            _managerInternal.SetDesktopWallpaper((IVirtualDesktop_Win11_24H2)objDesktop, path);
        }

        public void MoveWindowToDesktop(IntPtr hWnd, int desktopIndex)
        {
            if (hWnd == IntPtr.Zero) throw new ArgumentNullException();
            _managerInternal.GetDesktops(out var desktops);
            desktops.GetAt(desktopIndex, typeof(IVirtualDesktop_Win11_24H2).GUID, out var objDesktop);
            int processId;
            GetWindowThreadProcessId(hWnd, out processId);
            if (Process.GetCurrentProcess().Id == processId)
            {
                try
                {
                    _manager.MoveWindowToDesktop(hWnd, ((IVirtualDesktop_Win11_24H2)objDesktop).GetId());
                }
                catch
                {
                    _viewCollection.GetViewForHwnd(hWnd, out var view);
                    _managerInternal.MoveViewToDesktop(view, (IVirtualDesktop_Win11_24H2)objDesktop);
                }
            }
            else
            {
                _viewCollection.GetViewForHwnd(hWnd, out var view);
                try
                {
                    _managerInternal.MoveViewToDesktop(view, (IVirtualDesktop_Win11_24H2)objDesktop);
                }
                catch
                {
                    var mainHandle = Process.GetProcessById(processId).MainWindowHandle;
                    _viewCollection.GetViewForHwnd(mainHandle, out view);
                    _managerInternal.MoveViewToDesktop(view, (IVirtualDesktop_Win11_24H2)objDesktop);
                }
            }
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

        public void MoveActiveWindowToDesktop(int desktopIndex)
        {
            MoveWindowToDesktop(GetForegroundWindow(), desktopIndex);
        }

        public bool IsWindowOnDesktop(IntPtr hWnd, int desktopIndex)
        {
            if (hWnd == IntPtr.Zero) throw new ArgumentNullException();
            _managerInternal.GetDesktops(out var desktops);
            desktops.GetAt(desktopIndex, typeof(IVirtualDesktop_Win11_24H2).GUID, out var objDesktop);
            Guid id = _manager.GetWindowDesktopId(hWnd);
            return ((IVirtualDesktop_Win11_24H2)objDesktop).GetId() == id;
        }

        public bool IsWindowPinned(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero) throw new ArgumentNullException();
            _viewCollection.GetViewForHwnd(hWnd, out var view);
            return _pinnedApps.IsViewPinned(view);
        }

        public void PinWindow(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero) throw new ArgumentNullException();
            _viewCollection.GetViewForHwnd(hWnd, out var view);
            if (!_pinnedApps.IsViewPinned(view))
                _pinnedApps.PinView(view);
        }

        public void PinActiveWindow() => PinWindow(GetForegroundWindow());

        public void UnpinWindow(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero) throw new ArgumentNullException();
            _viewCollection.GetViewForHwnd(hWnd, out var view);
            if (_pinnedApps.IsViewPinned(view))
                _pinnedApps.UnpinView(view);
        }

        public void UnpinActiveWindow() => UnpinWindow(GetForegroundWindow());

        public bool IsApplicationPinned(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero) throw new ArgumentNullException();
            string appId = GetAppId(hWnd);
            return _pinnedApps.IsAppIdPinned(appId);
        }

        public void PinApplication(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero) throw new ArgumentNullException();
            string appId = GetAppId(hWnd);
            if (!_pinnedApps.IsAppIdPinned(appId))
                _pinnedApps.PinAppID(appId);
        }

        public void UnpinApplication(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero) throw new ArgumentNullException();
            string appId = GetAppId(hWnd);
            if (_pinnedApps.IsAppIdPinned(appId))
                _pinnedApps.UnpinAppID(appId);
        }

        private string GetAppId(IntPtr hWnd)
        {
            _viewCollection.GetViewForHwnd(hWnd, out var view);
            var getAppUserModelId = view.GetType().GetMethod("GetAppUserModelId");
            object[] parameters = new object[] { null };
            getAppUserModelId.Invoke(view, parameters);
            return parameters[0] as string;
        }

        public int GetLeftDesktopIndex(int index)
        {
            _managerInternal.GetDesktops(out var desktops);
            desktops.GetAt(index, typeof(IVirtualDesktop_Win11_24H2).GUID, out var objDesktop);
            int hr = _managerInternal.GetAdjacentDesktop((IVirtualDesktop_Win11_24H2)objDesktop, 3, out var leftDesktop); // 3 = LeftDirection
            if (hr == 0)
            {
                for (int i = 0; i < GetDesktopCount(); i++)
                {
                    desktops.GetAt(i, typeof(IVirtualDesktop_Win11_24H2).GUID, out var obj);
                    if (((IVirtualDesktop_Win11_24H2)obj).GetId() == leftDesktop.GetId())
                        return i;
                }
            }
            return -1;
        }

        public int GetRightDesktopIndex(int index)
        {
            _managerInternal.GetDesktops(out var desktops);
            desktops.GetAt(index, typeof(IVirtualDesktop_Win11_24H2).GUID, out var objDesktop);
            int hr = _managerInternal.GetAdjacentDesktop((IVirtualDesktop_Win11_24H2)objDesktop, 4, out var rightDesktop); // 4 = RightDirection
            if (hr == 0)
            {
                for (int i = 0; i < GetDesktopCount(); i++)
                {
                    desktops.GetAt(i, typeof(IVirtualDesktop_Win11_24H2).GUID, out var obj);
                    if (((IVirtualDesktop_Win11_24H2)obj).GetId() == rightDesktop.GetId())
                        return i;
                }
            }
            return -1;
        }
    }
} 