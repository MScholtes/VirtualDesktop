// Author: Markus Scholtes, 2023
// Version 1.16, 2023-09-17
// Version for Windows 10 1809 to 22H2
// Compile with:
// C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe VirtualDesktop.cs

using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

// set attributes
using System.Reflection;
[assembly:AssemblyTitle("Command line tool to manage virtual desktops")]
[assembly:AssemblyDescription("Command line tool to manage virtual desktops")]
[assembly:AssemblyConfiguration("")]
[assembly:AssemblyCompany("MS")]
[assembly:AssemblyProduct("VirtualDesktop")]
[assembly:AssemblyCopyright("� Markus Scholtes 2023")]
[assembly:AssemblyTrademark("")]
[assembly:AssemblyCulture("")]
[assembly:AssemblyVersion("1.16.0.0")]
[assembly:AssemblyFileVersion("1.16.0.0")]

// Based on http://stackoverflow.com/a/32417530, Windows 10 SDK, github project Grabacr07/VirtualDesktop and own research

namespace VirtualDesktop
{
	#region COM API
	internal static class Guids
	{
		public static readonly Guid CLSID_ImmersiveShell = new Guid("C2F03A33-21F5-47FA-B4BB-156362A2F239");
		public static readonly Guid CLSID_VirtualDesktopManagerInternal = new Guid("C5E0CDCA-7B6E-41B2-9FC4-D93975CC467B");
		public static readonly Guid CLSID_VirtualDesktopManager = new Guid("AA509086-5CA9-4C25-8F95-589D3C07B48A");
		public static readonly Guid CLSID_VirtualDesktopPinnedApps = new Guid("B5A399E7-1C87-46B8-88E9-FC5747B171BD");
	}

	[StructLayout(LayoutKind.Sequential)]
	internal struct Size
	{
		public int X;
		public int Y;
	}

	[StructLayout(LayoutKind.Sequential)]
	internal struct Rect
	{
		public int Left;
		public int Top;
		public int Right;
		public int Bottom;
	}

	internal enum APPLICATION_VIEW_CLOAK_TYPE : int
	{
		AVCT_NONE = 0,
		AVCT_DEFAULT = 1,
		AVCT_VIRTUAL_DESKTOP = 2
	}

	internal enum APPLICATION_VIEW_COMPATIBILITY_POLICY : int
	{
		AVCP_NONE = 0,
		AVCP_SMALL_SCREEN = 1,
		AVCP_TABLET_SMALL_SCREEN = 2,
		AVCP_VERY_SMALL_SCREEN = 3,
		AVCP_HIGH_SCALE_FACTOR = 4
	}

	[ComImport]
	[InterfaceType(ComInterfaceType.InterfaceIsIInspectable)]
	[Guid("372E1D3B-38D3-42E4-A15B-8AB2B178F513")]
	internal interface IApplicationView
	{
		int SetFocus();
		int SwitchTo();
		int TryInvokeBack(IntPtr /* IAsyncCallback* */ callback);
		int GetThumbnailWindow(out IntPtr hwnd);
		int GetMonitor(out IntPtr /* IImmersiveMonitor */ immersiveMonitor);
		int GetVisibility(out int visibility);
		int SetCloak(APPLICATION_VIEW_CLOAK_TYPE cloakType, int unknown);
		int GetPosition(ref Guid guid /* GUID for IApplicationViewPosition */, out IntPtr /* IApplicationViewPosition** */ position);
		int SetPosition(ref IntPtr /* IApplicationViewPosition* */ position);
		int InsertAfterWindow(IntPtr hwnd);
		int GetExtendedFramePosition(out Rect rect);
		int GetAppUserModelId([MarshalAs(UnmanagedType.LPWStr)] out string id);
		int SetAppUserModelId(string id);
		int IsEqualByAppUserModelId(string id, out int result);
		int GetViewState(out uint state);
		int SetViewState(uint state);
		int GetNeediness(out int neediness);
		int GetLastActivationTimestamp(out ulong timestamp);
		int SetLastActivationTimestamp(ulong timestamp);
		int GetVirtualDesktopId(out Guid guid);
		int SetVirtualDesktopId(ref Guid guid);
		int GetShowInSwitchers(out int flag);
		int SetShowInSwitchers(int flag);
		int GetScaleFactor(out int factor);
		int CanReceiveInput(out bool canReceiveInput);
		int GetCompatibilityPolicyType(out APPLICATION_VIEW_COMPATIBILITY_POLICY flags);
		int SetCompatibilityPolicyType(APPLICATION_VIEW_COMPATIBILITY_POLICY flags);
		int GetSizeConstraints(IntPtr /* IImmersiveMonitor* */ monitor, out Size size1, out Size size2);
		int GetSizeConstraintsForDpi(uint uint1, out Size size1, out Size size2);
		int SetSizeConstraintsForDpi(ref uint uint1, ref Size size1, ref Size size2);
		int OnMinSizePreferencesUpdated(IntPtr hwnd);
		int ApplyOperation(IntPtr /* IApplicationViewOperation* */ operation);
		int IsTray(out bool isTray);
		int IsInHighZOrderBand(out bool isInHighZOrderBand);
		int IsSplashScreenPresented(out bool isSplashScreenPresented);
		int Flash();
		int GetRootSwitchableOwner(out IApplicationView rootSwitchableOwner);
		int EnumerateOwnershipTree(out IObjectArray ownershipTree);
		int GetEnterpriseId([MarshalAs(UnmanagedType.LPWStr)] out string enterpriseId);
		int IsMirrored(out bool isMirrored);
		int Unknown1(out int unknown);
		int Unknown2(out int unknown);
		int Unknown3(out int unknown);
		int Unknown4(out int unknown);
		int Unknown5(out int unknown);
		int Unknown6(int unknown);
		int Unknown7();
		int Unknown8(out int unknown);
		int Unknown9(int unknown);
		int Unknown10(int unknownX, int unknownY);
		int Unknown11(int unknown);
		int Unknown12(out Size size1);
	}

	[ComImport]
	[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	[Guid("1841C6D7-4F9D-42C0-AF41-8747538F10E5")]
	internal interface IApplicationViewCollection
	{
		int GetViews(out IObjectArray array);
		int GetViewsByZOrder(out IObjectArray array);
		int GetViewsByAppUserModelId(string id, out IObjectArray array);
		int GetViewForHwnd(IntPtr hwnd, out IApplicationView view);
		int GetViewForApplication(object application, out IApplicationView view);
		int GetViewForAppUserModelId(string id, out IApplicationView view);
		int GetViewInFocus(out IntPtr view);
		int Unknown1(out IntPtr view);
		void RefreshCollection();
		int RegisterForApplicationViewChanges(object listener, out int cookie);
		int UnregisterForApplicationViewChanges(int cookie);
	}

	[ComImport]
	[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	[Guid("FF72FFDD-BE7E-43FC-9C03-AD81681E88E4")]
	internal interface IVirtualDesktop
	{
		bool IsViewVisible(IApplicationView view);
		Guid GetId();
	}

	[ComImport]
	[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	[Guid("F31574D6-B682-4CDC-BD56-1827860ABEC6")]
	internal interface IVirtualDesktopManagerInternal
	{
		int GetCount();
		void MoveViewToDesktop(IApplicationView view, IVirtualDesktop desktop);
		bool CanViewMoveDesktops(IApplicationView view);
		IVirtualDesktop GetCurrentDesktop();
		void GetDesktops(out IObjectArray desktops);
		[PreserveSig]
		int GetAdjacentDesktop(IVirtualDesktop from, int direction, out IVirtualDesktop desktop);
		void SwitchDesktop(IVirtualDesktop desktop);
		IVirtualDesktop CreateDesktop();
		void RemoveDesktop(IVirtualDesktop desktop, IVirtualDesktop fallback);
		IVirtualDesktop FindDesktop(ref Guid desktopid);
	}

	[ComImport]
	[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	[Guid("0F3A72B0-4566-487E-9A33-4ED302F6D6CE")]
	internal interface IVirtualDesktopManagerInternal2
	{
		int GetCount();
		void MoveViewToDesktop(IApplicationView view, IVirtualDesktop desktop);
		bool CanViewMoveDesktops(IApplicationView view);
		IVirtualDesktop GetCurrentDesktop();
		void GetDesktops(out IObjectArray desktops);
		[PreserveSig]
		int GetAdjacentDesktop(IVirtualDesktop from, int direction, out IVirtualDesktop desktop);
		void SwitchDesktop(IVirtualDesktop desktop);
		IVirtualDesktop CreateDesktop();
		void RemoveDesktop(IVirtualDesktop desktop, IVirtualDesktop fallback);
		IVirtualDesktop FindDesktop(ref Guid desktopid);
		void Unknown1(IVirtualDesktop desktop, out IntPtr unknown1, out IntPtr unknown2);
		void SetName(IVirtualDesktop desktop, [MarshalAs(UnmanagedType.HString)] string name);
	}

	[ComImport]
	[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	[Guid("A5CD92FF-29BE-454C-8D04-D82879FB3F1B")]
	internal interface IVirtualDesktopManager
	{
		bool IsWindowOnCurrentVirtualDesktop(IntPtr topLevelWindow);
		Guid GetWindowDesktopId(IntPtr topLevelWindow);
		void MoveWindowToDesktop(IntPtr topLevelWindow, ref Guid desktopId);
	}

	[ComImport]
	[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	[Guid("4CE81583-1E4C-4632-A621-07A53543148F")]
	internal interface IVirtualDesktopPinnedApps
	{
		bool IsAppIdPinned(string appId);
		void PinAppID(string appId);
		void UnpinAppID(string appId);
		bool IsViewPinned(IApplicationView applicationView);
		void PinView(IApplicationView applicationView);
		void UnpinView(IApplicationView applicationView);
	}

	[ComImport]
	[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	[Guid("92CA9DCD-5622-4BBA-A805-5E9F541BD8C9")]
	internal interface IObjectArray
	{
		void GetCount(out int count);
		void GetAt(int index, ref Guid iid, [MarshalAs(UnmanagedType.Interface)]out object obj);
	}

	[ComImport]
	[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	[Guid("6D5140C1-7436-11CE-8034-00AA006009FA")]
	internal interface IServiceProvider10
	{
		[return: MarshalAs(UnmanagedType.IUnknown)]
		object QueryService(ref Guid service, ref Guid riid);
	}
	#endregion

	#region COM wrapper
	internal static class DesktopManager
	{
		static DesktopManager()
		{
			var shell = (IServiceProvider10)Activator.CreateInstance(Type.GetTypeFromCLSID(Guids.CLSID_ImmersiveShell));
			VirtualDesktopManagerInternal = (IVirtualDesktopManagerInternal)shell.QueryService(Guids.CLSID_VirtualDesktopManagerInternal, typeof(IVirtualDesktopManagerInternal).GUID);
			try {
				VirtualDesktopManagerInternal2 = (IVirtualDesktopManagerInternal2)shell.QueryService(Guids.CLSID_VirtualDesktopManagerInternal, typeof(IVirtualDesktopManagerInternal2).GUID);
			}
			catch {
				VirtualDesktopManagerInternal2 = null;
			}
			VirtualDesktopManager = (IVirtualDesktopManager)Activator.CreateInstance(Type.GetTypeFromCLSID(Guids.CLSID_VirtualDesktopManager));
			ApplicationViewCollection = (IApplicationViewCollection)shell.QueryService(typeof(IApplicationViewCollection).GUID, typeof(IApplicationViewCollection).GUID);
			VirtualDesktopPinnedApps = (IVirtualDesktopPinnedApps)shell.QueryService(Guids.CLSID_VirtualDesktopPinnedApps, typeof(IVirtualDesktopPinnedApps).GUID);
		}

		internal static IVirtualDesktopManagerInternal VirtualDesktopManagerInternal;
		internal static IVirtualDesktopManagerInternal2 VirtualDesktopManagerInternal2;
		internal static IVirtualDesktopManager VirtualDesktopManager;
		internal static IApplicationViewCollection ApplicationViewCollection;
		internal static IVirtualDesktopPinnedApps VirtualDesktopPinnedApps;

		internal static IVirtualDesktop GetDesktop(int index)
		{	// get desktop with index
			int count = VirtualDesktopManagerInternal.GetCount();
			if (index < 0 || index >= count) throw new ArgumentOutOfRangeException("index");
			IObjectArray desktops;
			VirtualDesktopManagerInternal.GetDesktops(out desktops);
			object objdesktop;
			desktops.GetAt(index, typeof(IVirtualDesktop).GUID, out objdesktop);
			Marshal.ReleaseComObject(desktops);
			return (IVirtualDesktop)objdesktop;
		}

		internal static int GetDesktopIndex(IVirtualDesktop desktop)
		{ // get index of desktop
			int index = -1;
			Guid IdSearch = desktop.GetId();
			IObjectArray desktops;
			VirtualDesktopManagerInternal.GetDesktops(out desktops);
			object objdesktop;
			for (int i = 0; i < VirtualDesktopManagerInternal.GetCount(); i++)
			{
				desktops.GetAt(i, typeof(IVirtualDesktop).GUID, out objdesktop);
				if (IdSearch.CompareTo(((IVirtualDesktop)objdesktop).GetId()) == 0)
				{ index = i;
					break;
				}
			}
			Marshal.ReleaseComObject(desktops);
			return index;
		}

		internal static IApplicationView GetApplicationView(this IntPtr hWnd)
		{ // get application view to window handle
			IApplicationView view;
			ApplicationViewCollection.GetViewForHwnd(hWnd, out view);
			return view;
		}

		internal static string GetAppId(IntPtr hWnd)
		{ // get Application ID to window handle
			string appId;
			hWnd.GetApplicationView().GetAppUserModelId(out appId);
			return appId;
		}
	}
	#endregion

	#region public interface
	public class WindowInformation
	{ // stores window informations
		public string Title { get; set; }
		public int Handle { get; set; }
	}

	public class Desktop
	{
		// get process id to window handle
		[DllImport("user32.dll")]
		private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

		// get thread id of current process
		[DllImport("kernel32.dll")]
		static extern uint GetCurrentThreadId();

		// attach input to thread
		[DllImport("user32.dll")]
		static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

		// get handle of active window
		[DllImport("user32.dll")]
		private static extern IntPtr GetForegroundWindow();

		// try to set foreground window
		[DllImport("user32.dll")]
		[return: MarshalAs(UnmanagedType.Bool)]static extern bool SetForegroundWindow(IntPtr hWnd);

		// send message to window
		[DllImport("user32.dll")]
		static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
		private const int SW_MINIMIZE = 6;

		private static readonly Guid AppOnAllDesktops = new Guid("BB64D5B7-4DE3-4AB2-A87C-DB7601AEA7DC");
		private static readonly Guid WindowOnAllDesktops = new Guid("C2DDEA68-66F2-4CF9-8264-1BFD00FBBBAC");

		private IVirtualDesktop ivd;
		private Desktop(IVirtualDesktop desktop) { this.ivd = desktop; }

		public override int GetHashCode()
		{ // get hash
			return ivd.GetHashCode();
		}

		public override bool Equals(object obj)
		{ // compare with object
			var desk = obj as Desktop;
			return desk != null && object.ReferenceEquals(this.ivd, desk.ivd);
		}

		public static int Count
		{ // return the number of desktops
			get { return DesktopManager.VirtualDesktopManagerInternal.GetCount(); }
		}

		public static Desktop Current
		{ // returns current desktop
			get { return new Desktop(DesktopManager.VirtualDesktopManagerInternal.GetCurrentDesktop()); }
		}

		public static Desktop FromIndex(int index)
		{ // return desktop object from index (-> index = 0..Count-1)
			return new Desktop(DesktopManager.GetDesktop(index));
		}

		public static Desktop FromWindow(IntPtr hWnd)
		{ // return desktop object to desktop on which window <hWnd> is displayed
			if (hWnd == IntPtr.Zero) throw new ArgumentNullException();
			Guid id = DesktopManager.VirtualDesktopManager.GetWindowDesktopId(hWnd);
			if ((id.CompareTo(AppOnAllDesktops) == 0) || (id.CompareTo(WindowOnAllDesktops) == 0))
				return new Desktop(DesktopManager.VirtualDesktopManagerInternal.GetCurrentDesktop());
			else
				return new Desktop(DesktopManager.VirtualDesktopManagerInternal.FindDesktop(ref id));
		}

		public static int FromDesktop(Desktop desktop)
		{ // return index of desktop object or -1 if not found
			return DesktopManager.GetDesktopIndex(desktop.ivd);
		}

		public static string DesktopNameFromDesktop(Desktop desktop)
		{ // return name of desktop or "Desktop n" if it has no name
			Guid guid = desktop.ivd.GetId();

			// read desktop name in registry
			string desktopName = null;
			try {
				desktopName = (string)Microsoft.Win32.Registry.GetValue("HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\VirtualDesktops\\Desktops\\{" + guid.ToString() + "}", "Name", null);
			}
			catch { }

			// no name found, generate generic name
			if (string.IsNullOrEmpty(desktopName))
			{ // create name "Desktop n" (n = number starting with 1)
				desktopName = "Desktop " + (DesktopManager.GetDesktopIndex(desktop.ivd) + 1).ToString();
			}
			return desktopName;
		}

		public static string DesktopNameFromIndex(int index)
		{ // return name of desktop from index (-> index = 0..Count-1) or "Desktop n" if it has no name
			Guid guid = DesktopManager.GetDesktop(index).GetId();

			// read desktop name in registry
			string desktopName = null;
			try {
				desktopName = (string)Microsoft.Win32.Registry.GetValue("HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\VirtualDesktops\\Desktops\\{" + guid.ToString() + "}", "Name", null);
			}
			catch { }

			// no name found, generate generic name
			if (string.IsNullOrEmpty(desktopName))
			{ // create name "Desktop n" (n = number starting with 1)
				desktopName = "Desktop " + (index + 1).ToString();
			}
			return desktopName;
		}

		public static bool HasDesktopNameFromIndex(int index)
		{ // return true is desktop is named or false if it has no name
			Guid guid = DesktopManager.GetDesktop(index).GetId();

			// read desktop name in registry
			string desktopName = null;
			try {
				desktopName = (string)Microsoft.Win32.Registry.GetValue("HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\VirtualDesktops\\Desktops\\{" + guid.ToString() + "}", "Name", null);
			}
			catch { }

			// name found?
			if (string.IsNullOrEmpty(desktopName))
				return false;
			else
				return true;
		}

		public static int SearchDesktop(string partialName)
		{ // get index of desktop with partial name, return -1 if no desktop found
			int index = -1;

			for (int i = 0; i < DesktopManager.VirtualDesktopManagerInternal.GetCount(); i++)
			{ // loop through all virtual desktops and compare partial name to desktop name
				if (DesktopNameFromIndex(i).ToUpper().IndexOf(partialName.ToUpper()) >= 0)
				{ index = i;
					break;
				}
			}

			return index;
		}

		public static Desktop Create()
		{ // create a new desktop
			return new Desktop(DesktopManager.VirtualDesktopManagerInternal.CreateDesktop());
		}

		public void Remove(Desktop fallback = null)
		{ // destroy desktop and switch to <fallback>
			IVirtualDesktop fallbackdesktop;
			if (fallback == null)
			{ // if no fallback is given use desktop to the left except for desktop 0.
				Desktop dtToCheck = new Desktop(DesktopManager.GetDesktop(0));
				if (this.Equals(dtToCheck))
				{ // desktop 0: set fallback to second desktop (= "right" desktop)
					DesktopManager.VirtualDesktopManagerInternal.GetAdjacentDesktop(ivd, 4, out fallbackdesktop); // 4 = RightDirection
				}
				else
				{ // set fallback to "left" desktop
					DesktopManager.VirtualDesktopManagerInternal.GetAdjacentDesktop(ivd, 3, out fallbackdesktop); // 3 = LeftDirection
				}
			}
			else
				// set fallback desktop
				fallbackdesktop = fallback.ivd;

			DesktopManager.VirtualDesktopManagerInternal.RemoveDesktop(ivd, fallbackdesktop);
		}

		public static void RemoveAll()
		{ // remove all desktops but visible
			int desktopcount = DesktopManager.VirtualDesktopManagerInternal.GetCount();
			int desktopcurrent = DesktopManager.GetDesktopIndex(DesktopManager.VirtualDesktopManagerInternal.GetCurrentDesktop());
			
			if (desktopcurrent < desktopcount-1)
			{ // remove all desktops "right" from current
				for (int i = desktopcount-1; i > desktopcurrent; i--)
					DesktopManager.VirtualDesktopManagerInternal.RemoveDesktop(DesktopManager.GetDesktop(i), DesktopManager.VirtualDesktopManagerInternal.GetCurrentDesktop());
			}
			if (desktopcurrent > 0)
			{ // remove all desktops "left" from current
				for (int i = 0; i < desktopcurrent; i++)
					DesktopManager.VirtualDesktopManagerInternal.RemoveDesktop(DesktopManager.GetDesktop(0), DesktopManager.VirtualDesktopManagerInternal.GetCurrentDesktop());
			}
		}

		public void SetName(string Name)
		{ // set name for desktop, empty string removes name
			if (DesktopManager.VirtualDesktopManagerInternal2 != null)
			{ // only if interface to set name is present
				DesktopManager.VirtualDesktopManagerInternal2.SetName(this.ivd, Name);
			}
		}

		public bool IsVisible
		{ // return true if this desktop is the current displayed one
			get { return object.ReferenceEquals(ivd, DesktopManager.VirtualDesktopManagerInternal.GetCurrentDesktop()); }
		}

		public void MakeVisible()
		{ // make this desktop visible
			WindowInformation wi = FindWindow("Program Manager");

			// activate desktop to prevent flashing icons in taskbar
			int dummy;
			uint DesktopThreadId = GetWindowThreadProcessId(new IntPtr(wi.Handle), out dummy);
			uint ForegroundThreadId = GetWindowThreadProcessId(GetForegroundWindow(), out dummy);
			uint CurrentThreadId = GetCurrentThreadId();

			if ((DesktopThreadId != 0) && (ForegroundThreadId != 0) && (ForegroundThreadId != CurrentThreadId))
			{
				AttachThreadInput(DesktopThreadId, CurrentThreadId, true);
				AttachThreadInput(ForegroundThreadId, CurrentThreadId, true);
				SetForegroundWindow(new IntPtr(wi.Handle));
				AttachThreadInput(ForegroundThreadId, CurrentThreadId, false);
				AttachThreadInput(DesktopThreadId, CurrentThreadId, false);
			}

			DesktopManager.VirtualDesktopManagerInternal.SwitchDesktop(ivd);

			// direct desktop to give away focus
			ShowWindow(new IntPtr(wi.Handle), SW_MINIMIZE);
		}

		public Desktop Left
		{ // return desktop at the left of this one, null if none
			get
			{
				IVirtualDesktop desktop;
				int hr = DesktopManager.VirtualDesktopManagerInternal.GetAdjacentDesktop(ivd, 3, out desktop); // 3 = LeftDirection
				if (hr == 0)
					return new Desktop(desktop);
				else
					return null;
			}
		}

		public Desktop Right
		{ // return desktop at the right of this one, null if none
			get
			{
				IVirtualDesktop desktop;
				int hr = DesktopManager.VirtualDesktopManagerInternal.GetAdjacentDesktop(ivd, 4, out desktop); // 4 = RightDirection
				if (hr == 0)
					return new Desktop(desktop);
				else
					return null;
			}
		}

		public void MoveWindow(IntPtr hWnd)
		{ // move window to this desktop
			int processId;
			if (hWnd == IntPtr.Zero) throw new ArgumentNullException();
			GetWindowThreadProcessId(hWnd, out processId);

			if (System.Diagnostics.Process.GetCurrentProcess().Id == processId)
			{ // window of process
				try // the easy way (if we are owner)
				{
					DesktopManager.VirtualDesktopManager.MoveWindowToDesktop(hWnd, ivd.GetId());
				}
				catch // window of process, but we are not the owner
				{
					IApplicationView view;
					DesktopManager.ApplicationViewCollection.GetViewForHwnd(hWnd, out view);
					DesktopManager.VirtualDesktopManagerInternal.MoveViewToDesktop(view, ivd);
				}
			}
			else
			{ // window of other process
				IApplicationView view;
				DesktopManager.ApplicationViewCollection.GetViewForHwnd(hWnd, out view);
				try {
					DesktopManager.VirtualDesktopManagerInternal.MoveViewToDesktop(view, ivd);
				}
				catch
				{ // could not move active window, try main window (or whatever windows thinks is the main window)
					DesktopManager.ApplicationViewCollection.GetViewForHwnd(System.Diagnostics.Process.GetProcessById(processId).MainWindowHandle, out view);
					DesktopManager.VirtualDesktopManagerInternal.MoveViewToDesktop(view, ivd);
				}
			}
		}

		public void MoveActiveWindow()
		{ // move active window to this desktop
			MoveWindow(GetForegroundWindow());
		}

		public bool HasWindow(IntPtr hWnd)
		{ // return true if window is on this desktop
			if (hWnd == IntPtr.Zero) throw new ArgumentNullException();
			Guid id = DesktopManager.VirtualDesktopManager.GetWindowDesktopId(hWnd);
			if ((id.CompareTo(AppOnAllDesktops) == 0) || (id.CompareTo(WindowOnAllDesktops) == 0))
				return true;
			else
				return ivd.GetId() == id;
		}

		public static bool IsWindowPinned(IntPtr hWnd)
		{ // return true if window is pinned to all desktops
			if (hWnd == IntPtr.Zero) throw new ArgumentNullException();
			return DesktopManager.VirtualDesktopPinnedApps.IsViewPinned(hWnd.GetApplicationView());
		}

		public static void PinWindow(IntPtr hWnd)
		{ // pin window to all desktops
			if (hWnd == IntPtr.Zero) throw new ArgumentNullException();
			var view = hWnd.GetApplicationView();
			if (!DesktopManager.VirtualDesktopPinnedApps.IsViewPinned(view))
			{ // pin only if not already pinned
				DesktopManager.VirtualDesktopPinnedApps.PinView(view);
			}
		}

		public static void UnpinWindow(IntPtr hWnd)
		{ // unpin window from all desktops
			if (hWnd == IntPtr.Zero) throw new ArgumentNullException();
			var view = hWnd.GetApplicationView();
			if (DesktopManager.VirtualDesktopPinnedApps.IsViewPinned(view))
			{ // unpin only if not already unpinned
				DesktopManager.VirtualDesktopPinnedApps.UnpinView(view);
			}
		}

		public static bool IsApplicationPinned(IntPtr hWnd)
		{ // return true if application for window is pinned to all desktops
			if (hWnd == IntPtr.Zero) throw new ArgumentNullException();
			return DesktopManager.VirtualDesktopPinnedApps.IsAppIdPinned(DesktopManager.GetAppId(hWnd));
		}

		public static void PinApplication(IntPtr hWnd)
		{ // pin application for window to all desktops
			if (hWnd == IntPtr.Zero) throw new ArgumentNullException();
			string appId = DesktopManager.GetAppId(hWnd);
			if (!DesktopManager.VirtualDesktopPinnedApps.IsAppIdPinned(appId))
			{ // pin only if not already pinned
				DesktopManager.VirtualDesktopPinnedApps.PinAppID(appId);
			}
		}

		public static void UnpinApplication(IntPtr hWnd)
		{ // unpin application for window from all desktops
			if (hWnd == IntPtr.Zero) throw new ArgumentNullException();
			var view = hWnd.GetApplicationView();
			string appId = DesktopManager.GetAppId(hWnd);
			if (DesktopManager.VirtualDesktopPinnedApps.IsAppIdPinned(appId))
			{ // unpin only if pinned
				DesktopManager.VirtualDesktopPinnedApps.UnpinAppID(appId);
			}
		}

		// prepare callback function for window enumeration
		private delegate bool CallBackPtr(int hwnd, int lParam);
		private static CallBackPtr callBackPtr = Callback;
		// list of window informations
		private static List<WindowInformation> WindowInformationList = new List<WindowInformation>();

		// enumerate windows
		[DllImport("User32.dll", SetLastError = true)]
		[return: MarshalAs(UnmanagedType.Bool)]
		private static extern bool EnumWindows(CallBackPtr lpEnumFunc, IntPtr lParam);

		// get window title length
		[DllImport("User32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		private static extern int GetWindowTextLength(IntPtr hWnd);

		// get window title
		[DllImport("User32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

		// callback function for window enumeration
		private static bool Callback(int hWnd, int lparam)
		{
			int length = GetWindowTextLength((IntPtr)hWnd);
			if (length > 0)
			{
				StringBuilder sb = new StringBuilder(length + 1);
				if (GetWindowText((IntPtr)hWnd, sb, sb.Capacity) > 0)
				{ WindowInformationList.Add(new WindowInformation {Handle = hWnd, Title = sb.ToString()}); }
			}
			return true;
		}

		// get list of all windows with title
		public static List<WindowInformation> GetWindows()
		{
			WindowInformationList = new List<WindowInformation>();
			EnumWindows(callBackPtr, IntPtr.Zero);
			return WindowInformationList;
		}

		// find first window with string in title
		public static WindowInformation FindWindow(string WindowTitle)
		{
			WindowInformationList = new List<WindowInformation>();
			EnumWindows(callBackPtr, IntPtr.Zero);
			WindowInformation result = WindowInformationList.Find(x => x.Title.IndexOf(WindowTitle, StringComparison.OrdinalIgnoreCase) >= 0);
			return result;
		}
	}
	#endregion
}


namespace VDeskTool
{
	static class Program
	{
		static bool verbose = true;
		static bool breakonerror = true;
		static bool wrapdesktops = false;
		static int rc = 0;
		

		static int Main(string[] args)
		{
			if (args.Length == 0)
			{ // no arguments, show help screen
				HelpScreen();
				return -2;
			}

			var serializer = new System.Web.Script.Serialization.JavaScriptSerializer();

			foreach (string arg in args)
			{
				System.Text.RegularExpressions.GroupCollection groups = System.Text.RegularExpressions.Regex.Match(arg, @"^[-\/]?([^:=]+)[:=]?(.*)$").Groups;

				if (groups.Count != 3)
				{ // parameter error
					rc = -2;
				}
				else
				{ // reset return code if on error
					if (rc < 0) rc = 0;

					if (groups[2].Value == "")
					{ // parameter without value
						string upperToken = groups[1].Value.ToUpper();
						switch(upperToken)
						{
							case "HELP": // help screen
							case "H":
							case "?":
								HelpScreen();
								return 0;

							case "QUIET": // don't display messages
							case "Q":
								verbose = false;
								break;

							case "VERBOSE": // display messages
							case "V":
								Console.WriteLine("Verbose mode enabled");
								verbose = true;
								break;

							case "BREAK": // break on error
							case "B":
								if (verbose) Console.WriteLine("Break on error enabled");
								breakonerror = true;
								break;

							case "CONTINUE": // continue on error
							case "CO":
								if (verbose) Console.WriteLine("Break on error disabled");
								breakonerror = false;
								break;

							case "WRAP": // wrap desktops using "LEFT" or "RIGHT"
							case "W":
								if (verbose) Console.WriteLine("Wrapping desktops enabled");
								wrapdesktops = true;
								break;

							case "NOWRAP": // do not wrap desktops
							case "NW":
								if (verbose) Console.WriteLine("Wrapping desktop disabled");
								wrapdesktops = false;
								break;

							case "COUNT": // get count of desktops
							case "C":
								rc = VirtualDesktop.Desktop.Count;
								if (verbose) Console.WriteLine("Count of desktops: " + rc);
								break;

							case "LIST": // show list of desktops
							case "LI":
								int desktopCount = VirtualDesktop.Desktop.Count;
								int visibleDesktop = VirtualDesktop.Desktop.FromDesktop(VirtualDesktop.Desktop.Current);
								if (verbose)
								{
									Console.WriteLine("Virtual desktops:");
									Console.WriteLine("-----------------");
								}
								for (int i = 0; i < desktopCount; i++)
								{
									if (i != visibleDesktop)
										Console.WriteLine(VirtualDesktop.Desktop.DesktopNameFromIndex(i));
									else
										Console.WriteLine(VirtualDesktop.Desktop.DesktopNameFromIndex(i) + " (visible)");
								}
								if (verbose) Console.WriteLine("\nCount of desktops: " + desktopCount);
								break;
							case "JSON":

							    desktopCount = VirtualDesktop.Desktop.Count;
							    visibleDesktop = VirtualDesktop.Desktop.FromDesktop(VirtualDesktop.Desktop.Current);
								

								Console.Write("{");
								Console.Write("\"count\":"+desktopCount+",");
								Console.Write("\"desktops\":[");
								string comma = "";
								for (int i = 0; i < desktopCount; i++)
								{
									Console.Write(comma+"{");
									Console.Write("\"name\":"+serializer.Serialize(VirtualDesktop.Desktop.DesktopNameFromIndex(i))+",");
									Console.Write("\"visible\":");
									
									if (i != visibleDesktop)
										Console.Write("false,");
									else
										Console.Write("true,");

									Console.Write("\"wallpaper\":");

									//if (string.IsNullOrEmpty(VirtualDesktop.Desktop.DesktopWallpaperFromIndex(i)))
										Console.Write("null");
									//else
									//	Console.Write( serializer.Serialize(VirtualDesktop.Desktop.DesktopWallpaperFromIndex(i)) );
									Console.Write("}");
									comma=",";
								}
								Console.Write("]"); 
								Console.WriteLine("}");
								break;

							case "INTERACTIVE":
							case "INT":
										
								string argstr = "";
								bool echo = true;

								int lastDT = VirtualDesktop.Desktop.FromDesktop(VirtualDesktop.Desktop.Current);

								verbose = false;
								while (true)
								{
								 
									if (Console.KeyAvailable )
									{
										var cki = Console.ReadKey(true);
										if (cki.Key == ConsoleKey.Escape) {
											break;
										} 

										if (cki.Key == ConsoleKey.Backspace) {
											 
											if (argstr != "" ) {
												argstr = argstr.Substring(0,argstr.Length-1);
											}  
											if (echo) {
												Console.Write(cki.KeyChar);
												Console.Write(" ");
												Console.Write(cki.KeyChar);
											}
											continue;
										}

										 
										if (cki.Key == ConsoleKey.Enter) {

											if (argstr == "") {
												continue;
											}
											if (echo) {
												Console.WriteLine("");
											}

											string[] splits = argstr.Split(' ');
											argstr = "";
											string token = splits[0].ToUpper();
											if (token.Length>1 && token.Substring(0,1)=="/") {
												token = token.Substring(1);
											}
											switch (token) {												
												case "INTERACTIVE":// prevent recursion
												case "INT":// prevent recursion

												case "BREAK":
												case "B":
												case "CONTINUE":
												case "CO":
												case "WAITDESKTOPCHANGE":
												case "WDC": 
													//Console.WriteLine();

													
													Console.WriteLine("\n{\"error\":"+  serializer.Serialize("Invalid Command:"+splits[0])+"}");
													continue;

												case "NAMES":

													int dtc = VirtualDesktop.Desktop.Count;
													Console.Write("[");
													string cma = "";
													for (int i = 0; i < dtc; i++)
													{
													    Console.Write(cma+serializer.Serialize(VirtualDesktop.Desktop.DesktopNameFromIndex(i)));
														cma=",";
													}
													Console.WriteLine("]");
												    continue;

												case "LEFT":
												case "P":
												case "PREVIOUS":
												   
												   if (lastDT==0) {
													   token="GCD";
												   } else {
													   token="L";													   
												   }
												   splits=token.Split(' ');
												   lastDT=-1;
												   break;

												case "L":

													if (lastDT==0) {
													   token="GCD";
													   splits=token.Split(' ');
												   }
												   lastDT=-1;
												   break;

												case "RIGHT":
												case "NEXT":
												   
												   if (lastDT==VirtualDesktop.Desktop.Count-1) {
													   token="GCD";
												   } else {
													   token="RI";													   
												   }
												   splits=token.Split(' ');
												   lastDT=-1;
												   break;

												case "RI":
												
													 if (lastDT==VirtualDesktop.Desktop.Count-1) {
													   token="GCD";
													   splits=token.Split(' ');
												    }
												   lastDT=-1;
												   break;

												case "GCD":
												case "GETCURRENTDESKTOP":
												   
												    lastDT=-1;// force a reporting of the desktop

													break;

											}
											
											if ( ((token.Length>=7) && (token.Substring(0,7)=="SWITCH:")) || 
											     ((token.Length>=2) && (token.Substring(0,2)=="S:")) ) {
												    lastDT=-1;// force a reporting of the desktop, even if the switched to desktop is the same as the current desktop
											}
	
											Main(splits);

											switch (token) {	
												case "NEW":
												case "N":
											 	Main(("S:"+rc).Split(' '));
											     //Console.WriteLine("\n{\"visibleIndex\":"+rc+",\"visible\":"+serializer.Serialize(VirtualDesktop.Desktop.DesktopNameFromIndex(rc))+"}");
												break;
											}

											 
										} else {
											argstr += cki.KeyChar;
											if (echo) {
												Console.Write(cki.KeyChar);
											}
											continue;
										}
										
									} else {
										int thisDT = VirtualDesktop.Desktop.FromDesktop(VirtualDesktop.Desktop.Current);
										if (lastDT != thisDT) {
											Console.WriteLine("\n{\"visibleIndex\":"+thisDT+",\"visible\":"+serializer.Serialize(VirtualDesktop.Desktop.DesktopNameFromIndex(thisDT))+"}");
											lastDT = thisDT;
										}

									}
								}


							break;


							case "WAITDESKTOPCHANGE": // wait for desktop to change
							case "WDC":
							case "GETCURRENTDESKTOP": // get number of current desktop and display desktop name
							case "GCD":
								rc = VirtualDesktop.Desktop.FromDesktop(VirtualDesktop.Desktop.Current);
								switch(upperToken) {
									case "WAITDESKTOPCHANGE":
									case "WDC":
									 
									if (verbose) Console.WriteLine("Waiting for desktop to change from: '" + VirtualDesktop.Desktop.DesktopNameFromDesktop(VirtualDesktop.Desktop.Current) + "' (desktop number " + rc + ")");
									int was=rc;
									while (rc == was) {
										System.Threading.Thread.Sleep(5);
										rc = VirtualDesktop.Desktop.FromDesktop(VirtualDesktop.Desktop.Current);
									}
									break;						
								}
								if (verbose) Console.WriteLine("Current desktop: '" + VirtualDesktop.Desktop.DesktopNameFromDesktop(VirtualDesktop.Desktop.Current) + "' (desktop number " + rc + ")");
								break;

							case "ISVISIBLE": // is desktop in rc visible?
							case "IV":
								if ((rc >= 0) && (rc < VirtualDesktop.Desktop.Count))
								{ // check if parameter is in range of active desktops
									if (VirtualDesktop.Desktop.FromIndex(rc).IsVisible)
									{
										if (verbose) Console.WriteLine("Virtual desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(rc) + "' (desktop number " + rc.ToString() + ") is visible");
										rc = 0;
									}
									else
									{
										if (verbose) Console.WriteLine("Virtual desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(rc) + "' (desktop number " + rc.ToString() + ") is not visible");
										rc = 1;
									}
								}
								else
									rc = -1;
								break;

							case "SWITCH": // switch to desktop in rc
							case "S":
								if (verbose) Console.Write("Switching to virtual desktop number " + rc.ToString());
								try
								{ // activate virtual desktop rc
									VirtualDesktop.Desktop.FromIndex(rc).MakeVisible();
									if (verbose) Console.WriteLine(", desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(rc) + "' is active now");
								}
								catch
								{ // error while activating
									if (verbose) Console.WriteLine();
									rc = -1;
								}
								break;

							case "LEFT": // switch to desktop to the left
							case "L":
								if (verbose) Console.Write("Switching to left virtual desktop");
								try
								{ // activate virtual desktop to the left
									if (wrapdesktops && (VirtualDesktop.Desktop.FromDesktop(VirtualDesktop.Desktop.Current) == 0))
										VirtualDesktop.Desktop.FromIndex(VirtualDesktop.Desktop.Count - 1).MakeVisible();
									else
										VirtualDesktop.Desktop.Current.Left.MakeVisible();
									rc = VirtualDesktop.Desktop.FromDesktop(VirtualDesktop.Desktop.Current);
									if (verbose) Console.WriteLine(", desktop number " + rc.ToString() + " ('" + VirtualDesktop.Desktop.DesktopNameFromIndex(rc) + "') is active now");
								}
								catch
								{ // error while activating
									if (verbose) Console.WriteLine();
									rc = -1;
								}
								break;

							case "RIGHT": // switch to desktop to the right
							case "RI":
								if (verbose) Console.Write("Switching to right virtual desktop");
								try
								{ // activate virtual desktop to the right
									if (wrapdesktops && (VirtualDesktop.Desktop.FromDesktop(VirtualDesktop.Desktop.Current) == VirtualDesktop.Desktop.Count - 1))
										VirtualDesktop.Desktop.FromIndex(0).MakeVisible();
									else
										VirtualDesktop.Desktop.Current.Right.MakeVisible();
									rc = VirtualDesktop.Desktop.FromDesktop(VirtualDesktop.Desktop.Current);
									if (verbose) Console.WriteLine(", desktop number " + rc.ToString() + " ('" + VirtualDesktop.Desktop.DesktopNameFromIndex(rc) + "') is active now");
								}
								catch
								{ // error while activating
									if (verbose) Console.WriteLine();
									rc = -1;
								}
								break;

							case "NEW": // create new desktop
							case "N":
								if (verbose) Console.Write("Creating virtual desktop");
								try
								{ // create virtual desktop, number is stored in rc
									rc = VirtualDesktop.Desktop.FromDesktop(VirtualDesktop.Desktop.Create());
									if (verbose) Console.WriteLine(" number " + rc.ToString());
								}
								catch
								{ // error while creating
									Console.WriteLine();
									rc = -1;
								}
								break;

							case "NAME": // removing name of desktop in rc
							case "NA":
									try
									{ // remove desktop name
										VirtualDesktop.Desktop.FromIndex(rc).SetName("");
										if (verbose) Console.WriteLine("Removed name of desktop number " + rc.ToString());
									}
									catch
									{ // error while removing name
										if (verbose) Console.WriteLine("Error removing desktop name");
										rc = -1;
									}
								break;

							case "REMOVE": // remove desktop in rc
							case "R":
								if (verbose)
								{
									Console.Write("Removing virtual desktop number " + rc.ToString());
									if ((rc >= 0) && (rc < VirtualDesktop.Desktop.Count)) Console.WriteLine(" (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(rc) + "')");
								}
								try
								{ // remove virtual desktop rc
									VirtualDesktop.Desktop.FromIndex(rc).Remove();
								}
								catch
								{ // error while removing
									Console.WriteLine();
									rc = -1;
								}
								break;

							case "REMOVEALL": // remove all virtual desktops but visible
							case "RA":
								Console.WriteLine("Removing all virtual desktops but visible");
								try
								{ // remove all virtual desktops but visible
									VirtualDesktop.Desktop.RemoveAll();
								}
								catch
								{ // error while removing
									rc = -1;
								}
								break;

							case "MOVEACTIVEWINDOW": // move active window to desktop in rc
							case "MAW":
								if (verbose) Console.WriteLine("Moving active window to virtual desktop number " + rc.ToString());
								try
								{ // move active window
									VirtualDesktop.Desktop.FromIndex(rc).MoveActiveWindow();
									if (verbose) Console.WriteLine("Active window moved to desktop number " + rc.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(rc) + "')");
								}
								catch
								{ // error while moving
									if (verbose) Console.WriteLine("No active window or move failed");
									rc = -1;
								}
								break;

							case "LISTWINDOWSONDESKTOP": // list window handles of windows shown on desktop in rc
							case "LWOD":
								if (verbose)
								{
									Console.Write("Listing handles of windows on virtual desktop number " + rc.ToString());
									if ((rc >= 0) && (rc < VirtualDesktop.Desktop.Count)) Console.WriteLine(" (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(rc) + "')");
								}
								try
								{ // list window handles on desktop rc
									ListWindowsOnDesktop(rc);
								}
								catch
								{ // error while listing
									Console.WriteLine();
									rc = -1;
								}
								break;

							case "CLOSEWINDOWSONDESKTOP": // close windows shown on desktop in rc
							case "CWOD":
								if (verbose)
								{
									Console.Write("Closing windows on virtual desktop number " + rc.ToString());
									if ((rc >= 0) && (rc < VirtualDesktop.Desktop.Count)) Console.WriteLine(" (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(rc) + "')");
								}
								try
								{ // close windows on desktop rc
									CloseWindowsOnDesktop(rc);
								}
								catch
								{ // error while closing
									Console.WriteLine();
									rc = -1;
								}
								break;

							case "WAITKEY": // wait for keypress
							case "WK":
								if (verbose) Console.WriteLine("Press a key");
								Console.ReadKey(true);
								break;

							default:
								rc = -2;
								break;
						}
					}
					else
					{	// parameter with value
						int iParam;

						switch(groups[1].Value.ToUpper())
						{
							case "GETDESKTOP": // get desktop number
							case "GD":
								if (int.TryParse(groups[2].Value, out iParam))
								{ // parameter is an integer, use as desktop number
									if ((iParam >= 0) && (iParam < VirtualDesktop.Desktop.Count))
									{ // check if parameter is in range of active desktops
										if (verbose) Console.WriteLine("Virtual desktop number " + iParam.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(iParam) + "') selected");
										rc = iParam;
									}
									else
										rc = -1;
								}
								else
								{ // parameter is a string, search as part of desktop name
									iParam = VirtualDesktop.Desktop.SearchDesktop(groups[2].Value);
									if (iParam >= 0)
									{ // desktop found
										if (verbose) Console.WriteLine("Virtual desktop number " + iParam.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(iParam) + "') selected");
										rc = iParam;
									}
									else
									{ // no desktop found
										if ((groups[2].Value.ToUpper() == "LAST") || (groups[2].Value.ToUpper() == "*LAST*"))
										{ // last desktop
											iParam = VirtualDesktop.Desktop.Count-1;
											if (verbose) Console.WriteLine("Virtual desktop number " + iParam.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(iParam) + "') selected");
											rc = iParam;
										}
										else
										{ // no desktop found
											if (verbose) Console.WriteLine("Could not find virtual desktop with name containing '" + groups[2].Value + "'");
											rc = -2;
										}
									}
								}
								break;

							case "ISVISIBLE": // is desktop visible?
							case "IV":
								if (int.TryParse(groups[2].Value, out iParam))
								{ // parameter is an integer, use as desktop number
									if ((iParam >= 0) && (iParam < VirtualDesktop.Desktop.Count))
									{ // check if parameter is in range of active desktops
										if (VirtualDesktop.Desktop.FromIndex(iParam).IsVisible)
										{
											if (verbose) Console.WriteLine("Virtual desktop number " + iParam.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(iParam) + "') is visible");
											rc = 0;
										}
										else
										{
											if (verbose) Console.WriteLine("Virtual desktop number " + iParam.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(iParam) + "') is not visible");
											rc = 1;
										}
									}
									else
										rc = -1;
								}
								else
								{ // parameter is a string, search as part of desktop name
									iParam = VirtualDesktop.Desktop.SearchDesktop(groups[2].Value);
									if (iParam >= 0)
									{ // desktop found
										if (VirtualDesktop.Desktop.FromIndex(iParam).IsVisible)
										{
											if (verbose) Console.WriteLine("Virtual desktop number " + iParam.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(iParam) + "') is visible");
											rc = 0;
										}
										else
										{
											if (verbose) Console.WriteLine("Virtual desktop number " + iParam.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(iParam) + "') is not visible");
											rc = 1;
										}
									}
									else
									{ // no desktop found
										if ((groups[2].Value.ToUpper() == "LAST") || (groups[2].Value.ToUpper() == "*LAST*"))
										{ // last desktop
											iParam = VirtualDesktop.Desktop.Count-1;
											if (VirtualDesktop.Desktop.FromIndex(iParam).IsVisible)
											{
												if (verbose) Console.WriteLine("Virtual desktop number " + iParam.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(iParam) + "') is visible");
												rc = 0;
											}
											else
											{
												if (verbose) Console.WriteLine("Virtual desktop number " + iParam.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(iParam) + "') is not visible");
												rc = 1;
											}
										}
										else
										{ // no desktop found
											if (verbose) Console.WriteLine("Could not find virtual desktop with name containing '" + groups[2].Value + "'");
											rc = -2;
										}
									}
								}
								break;

							case "NAME": // set name of desktop in rc
							case "NA":
									try
									{ // set desktop name
										VirtualDesktop.Desktop.FromIndex(rc).SetName(groups[2].Value);
										if (verbose) Console.WriteLine("Set name of desktop number " + rc.ToString() + " to '" + VirtualDesktop.Desktop.DesktopNameFromIndex(rc) + "'");
									}
									catch
									{ // error while setting name
										if (verbose) Console.WriteLine("Error setting desktop name to '" + groups[2].Value + "'");
										rc = -1;
									}
								break;

							case "SWITCH": // switch to desktop
							case "S":
								if (int.TryParse(groups[2].Value, out iParam))
								{ // parameter is an integer, use as desktop number
									if ((iParam >= 0) && (iParam < VirtualDesktop.Desktop.Count))
									{ // check if parameter is in range of active desktops
										if (verbose) Console.WriteLine("Switching to virtual desktop number " + iParam.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(iParam) + "')");
										rc = iParam;
										try
										{ // activate virtual desktop iParam
											VirtualDesktop.Desktop.FromIndex(iParam).MakeVisible();
										}
										catch
										{ // error while activating
											rc = -1;
										}
									}
									else
										rc = -1;
								}
								else
								{ // parameter is a string, search as part of desktop name
									iParam = VirtualDesktop.Desktop.SearchDesktop(groups[2].Value);
									if (iParam >= 0)
									{ // desktop found
										if (verbose) Console.WriteLine("Switching to virtual desktop number " + iParam.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(iParam) + "')");
										rc = iParam;
										try
										{ // activate virtual desktop iParam
											VirtualDesktop.Desktop.FromIndex(iParam).MakeVisible();
										}
										catch
										{ // error while activating
											rc = -1;
										}
									}
									else
									{ // no desktop found
										if ((groups[2].Value.ToUpper() == "LAST") || (groups[2].Value.ToUpper() == "*LAST*"))
										{ // last desktop
											iParam = VirtualDesktop.Desktop.Count-1;
											if (verbose) Console.WriteLine("Switching to virtual desktop number " + iParam.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(iParam) + "')");
											rc = iParam;
											try
											{ // activate virtual desktop iParam
												VirtualDesktop.Desktop.FromIndex(iParam).MakeVisible();
											}
											catch
											{ // error while activating
												rc = -1;
											}
										}
										else
										{ // no desktop found
											if (verbose) Console.WriteLine("Could not find virtual desktop with name containing '" + groups[2].Value + "'");
											rc = -2;
										}
									}
								}
								break;

							case "REMOVE": // remove desktop
							case "R":
								if (int.TryParse(groups[2].Value, out iParam))
								{ // parameter is an integer, use as desktop number
									if ((iParam >= 0) && (iParam < VirtualDesktop.Desktop.Count))
									{ // check if parameter is in range of active desktops
										if (verbose) Console.WriteLine("Removing virtual desktop number " + iParam.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(iParam) + "')");
										rc = iParam;
										try
										{ // remove virtual desktop iParam
											VirtualDesktop.Desktop.FromIndex(iParam).Remove();
										}
										catch
										{ // error while removing
											rc = -1;
										}
									}
									else
										rc = -1;
								}
								else
								{ // parameter is a string, search as part of desktop name
									iParam = VirtualDesktop.Desktop.SearchDesktop(groups[2].Value);
									if (iParam >= 0)
									{ // desktop found
										if (verbose) Console.WriteLine("Removing virtual desktop number " + iParam.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(iParam) + "')");
										rc = iParam;
										try
										{ // remove virtual desktop iParam
											VirtualDesktop.Desktop.FromIndex(iParam).Remove();
										}
										catch
										{ // error while removing
											rc = -1;
										}
									}
									else
									{ // no desktop found
										if ((groups[2].Value.ToUpper() == "LAST") || (groups[2].Value.ToUpper() == "*LAST*"))
										{ // last desktop
											iParam = VirtualDesktop.Desktop.Count-1;
											if (verbose) Console.WriteLine("Removing virtual desktop number " + iParam.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(iParam) + "')");
											rc = iParam;
											try
											{ // remove virtual desktop iParam
												VirtualDesktop.Desktop.FromIndex(iParam).Remove();
											}
											catch
											{ // error while removing
												rc = -1;
											}
										}
										else
										{ // no desktop found
											if (verbose) Console.WriteLine("Could not find virtual desktop with name containing '" + groups[2].Value + "'");
											rc = -2;
										}
									}
								}
								break;

							case "SWAPDESKTOP": // swap desktops
							case "SD":
								if (int.TryParse(groups[2].Value, out iParam))
								{ // parameter is an integer, use as desktop number
									if ((iParam >= 0) && (iParam < VirtualDesktop.Desktop.Count) && (rc != iParam))
									{ // check if parameter is in range of active desktops
										if (verbose) Console.WriteLine("Swapping virtual desktops number " + rc.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(rc) + "') and number " + iParam.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(iParam) + "')");
										try
										{ // swap virtual desktops rc and iParam
											SwapDesktops(rc, iParam);
											rc = iParam;
										}
										catch
										{ // error while swapping
											rc = -1;
										}
									}
									else
										rc = -1;
								}
								else
								{ // parameter is a string, search as part of desktop name
									iParam = VirtualDesktop.Desktop.SearchDesktop(groups[2].Value);
									if ((iParam >= 0) && (rc != iParam))
									{ // desktop found
										if (verbose) Console.WriteLine("Swapping virtual desktops number " + rc.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(rc) + "') and number " + iParam.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(iParam) + "')");
										try
										{ // swap virtual desktops rc and iParam
											SwapDesktops(rc, iParam);
											rc = iParam;
										}
										catch
										{ // error while swapping
											rc = -1;
										}
									}
									else
									{ // no desktop found or source and target the same
										if (rc == iParam)
										{
											if (verbose) Console.WriteLine("Cannot swap virtual desktop with itself");
											rc = -2;
										}
										else
										{ // no desktop found
											if ((groups[2].Value.ToUpper() == "LAST") || (groups[2].Value.ToUpper() == "*LAST*"))
											{ // last desktop
												iParam = VirtualDesktop.Desktop.Count-1;
												if (rc == iParam)
												{
													if (verbose) Console.WriteLine("Cannot swap virtual desktop with itself");
													rc = -2;
												}
												else
												{
													if (verbose) Console.WriteLine("Swapping virtual desktops number " + rc.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(rc) + "') and number " + iParam.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(iParam) + "')");
													try
													{ // swap virtual desktops rc and iParam
														SwapDesktops(rc, iParam);
														rc = iParam;
													}
													catch
													{ // error while swapping
														rc = -1;
													}
												}
											}
											else
											{ // no desktop found
												if (verbose) Console.WriteLine("Could not find virtual desktop with name containing '" + groups[2].Value + "'");
												rc = -2;
											}
										}
									}
								}
								break;

							case "INSERTDESKTOP": // insert desktop
							case "ID":
								if (int.TryParse(groups[2].Value, out iParam))
								{ // parameter is an integer, use as desktop number
									if ((iParam >= 0) && (iParam < VirtualDesktop.Desktop.Count) && (rc != iParam))
									{ // check if parameter is in range of active desktops
										if (verbose) Console.WriteLine("Inserting virtual desktop number " + iParam.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(iParam) + "') before desktop number " + rc.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(rc) + "') or vice versa");
										try
										{ // insert virtual desktop iParam before rc
											InsertDesktop(rc, iParam);
											rc = iParam;
										}
										catch
										{ // error while inserting
											rc = -1;
										}
									}
									else
										rc = -1;
								}
								else
								{ // parameter is a string, search as part of desktop name
									iParam = VirtualDesktop.Desktop.SearchDesktop(groups[2].Value);
									if ((iParam >= 0) && (rc != iParam))
									{ // desktop found
										if (verbose) Console.WriteLine("Inserting virtual desktop number " + iParam.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(iParam) + "') before desktop number " + rc.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(rc) + "') or vice versa");
										try
										{ // insert virtual desktop iParam before rc
											InsertDesktop(rc, iParam);
											rc = iParam;
										}
										catch
										{ // error while inserting
											rc = -1;
										}
									}
									else
									{ // no desktop found or source and target the same
										if (rc == iParam)
										{
											if (verbose) Console.WriteLine("Cannot insert virtual desktop before itself");
											rc = -2;
										}
										else
										{ // no desktop found
											if ((groups[2].Value.ToUpper() == "LAST") || (groups[2].Value.ToUpper() == "*LAST*"))
											{ // last desktop
												iParam = VirtualDesktop.Desktop.Count-1;
												if (rc == iParam)
												{
													if (verbose) Console.WriteLine("Cannot insert virtual desktop before itself");
													rc = -2;
												}
												else
												{
													if (verbose) Console.WriteLine("Inserting virtual desktop number " + iParam.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(iParam) + "') before desktop number " + rc.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(rc) + "') or vice versa");
													try
													{ // insert virtual desktop iParam before rc
														InsertDesktop(rc, iParam);
														rc = iParam;
													}
													catch
													{ // error while inserting
														rc = -1;
													}
												}
											}
											else
											{ // no desktop found
												if (verbose) Console.WriteLine("Could not find virtual desktop with name containing '" + groups[2].Value + "'");
												rc = -2;
											}
										}
									}
								}
								break;

							case "MoveWindowsToDesktop": // move windows of desktop in pipeline to desktop
							case "MWTD":
								if (int.TryParse(groups[2].Value, out iParam))
								{ // parameter is an integer, use as desktop number
									if ((iParam >= 0) && (iParam < VirtualDesktop.Desktop.Count) && (rc != iParam))
									{ // check if parameter is in range of active desktops
										if (verbose) Console.WriteLine("Moving windows on virtual desktop number " + rc.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(rc) + "') to desktop number " + iParam.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(iParam) + "')");
										try
										{ // move windows on desktop rc to desktop iParam
											MoveWindowsToDesktop(rc, iParam);
											rc = iParam;
										}
										catch
										{ // error while moving
											rc = -1;
										}
									}
									else
										rc = -1;
								}
								else
								{ // parameter is a string, search as part of desktop name
									iParam = VirtualDesktop.Desktop.SearchDesktop(groups[2].Value);
									if ((iParam >= 0) && (rc != iParam))
									{ // desktop found
										if (verbose) Console.WriteLine("Moving windows on virtual desktop number " + rc.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(rc) + "') to desktop number " + iParam.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(iParam) + "')");
										try
										{ // move windows on desktop rc to desktop iParam
											MoveWindowsToDesktop(rc, iParam);
											rc = iParam;
										}
										catch
										{ // error while moving
											rc = -1;
										}
									}
									else
									{ // no desktop found or source and target the same
										if (rc == iParam)
										{
											if (verbose) Console.WriteLine("Cannot move to same virtual desktop");
											rc = -2;
										}
										else
										{ // no desktop found
											if ((groups[2].Value.ToUpper() == "LAST") || (groups[2].Value.ToUpper() == "*LAST*"))
											{ // last desktop
												iParam = VirtualDesktop.Desktop.Count-1;
												if (rc == iParam)
												{
													if (verbose) Console.WriteLine("Cannot move to same virtual desktop");
													rc = -2;
												}
												else
												{
													if (verbose) Console.WriteLine("Moving windows on virtual desktop number " + rc.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(rc) + "') to desktop number " + iParam.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(iParam) + "')");
													try
													{ // move windows on desktop rc to desktop iParam
														MoveWindowsToDesktop(rc, iParam);
														rc = iParam;
													}
													catch
													{ // error while moving
														rc = -1;
													}
												}
											}
											else
											{ // no desktop found
												if (verbose) Console.WriteLine("Could not find virtual desktop with name containing '" + groups[2].Value + "'");
												rc = -2;
											}
										}
									}
								}
								break;

							case "LISTWINDOWSONDESKTOP": // list window handles of windows shown on desktop
							case "LWOD":
								if (int.TryParse(groups[2].Value, out iParam))
								{ // parameter is an integer, use as desktop number
									if ((iParam >= 0) && (iParam < VirtualDesktop.Desktop.Count))
									{ // check if parameter is in range of active desktops
										if (verbose) Console.WriteLine("Listing handles of windows on virtual desktop number " + iParam.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(iParam) + "')");
										try
										{ // list window handles on desktop iParam
											ListWindowsOnDesktop(iParam);
											rc = iParam;
										}
										catch
										{ // error while listing
											rc = -1;
										}
									}
									else
										rc = -1;
								}
								else
								{ // parameter is a string, search as part of desktop name
									iParam = VirtualDesktop.Desktop.SearchDesktop(groups[2].Value);
									if (iParam >= 0)
									{ // desktop found
										if (verbose) Console.WriteLine("Listing window handles of windows on virtual desktop number " + iParam.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(iParam) + "')");
										try
										{ // list window handles on desktop iParam
											ListWindowsOnDesktop(iParam);
											rc = iParam;
										}
										catch
										{ // error while listing
											rc = -1;
										}
									}
									else
									{ // no desktop found
										if ((groups[2].Value.ToUpper() == "LAST") || (groups[2].Value.ToUpper() == "*LAST*"))
										{ // last desktop
											iParam = VirtualDesktop.Desktop.Count-1;
											if (verbose) Console.WriteLine("Listing window handles of windows on virtual desktop number " + iParam.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(iParam) + "')");
											try
											{ // list window handles on desktop iParam
												ListWindowsOnDesktop(iParam);
												rc = iParam;
											}
											catch
											{ // error while listing
												rc = -1;
											}
										}
										else
										{ // no desktop found
											if (verbose) Console.WriteLine("Could not find virtual desktop with name containing '" + groups[2].Value + "'");
											rc = -2;
										}
									}
								}
								break;

							case "CLOSEWINDOWSONDESKTOP": // close windows shown on desktop
							case "CWOD":
								if (int.TryParse(groups[2].Value, out iParam))
								{ // parameter is an integer, use as desktop number
									if ((iParam >= 0) && (iParam < VirtualDesktop.Desktop.Count))
									{ // check if parameter is in range of active desktops
										if (verbose) Console.WriteLine("Closing windows on virtual desktop number " + iParam.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(iParam) + "')");
										try
										{ // close windows on desktop iParam
											CloseWindowsOnDesktop(iParam);
											rc = iParam;
										}
										catch
										{ // error while closing
											rc = -1;
										}
									}
									else
										rc = -1;
								}
								else
								{ // parameter is a string, search as part of desktop name
									iParam = VirtualDesktop.Desktop.SearchDesktop(groups[2].Value);
									if (iParam >= 0)
									{ // desktop found
										if (verbose) Console.WriteLine("Closing windows on virtual desktop number " + iParam.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(iParam) + "')");
										try
										{ // close windows on desktop iParam
											CloseWindowsOnDesktop(iParam);
											rc = iParam;
										}
										catch
										{ // error while closing
											rc = -1;
										}
									}
									else
									{ // no desktop found
										if ((groups[2].Value.ToUpper() == "LAST") || (groups[2].Value.ToUpper() == "*LAST*"))
										{ // last desktop
											iParam = VirtualDesktop.Desktop.Count-1;
											if (verbose) Console.WriteLine("Closing windows on virtual desktop number " + iParam.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(iParam) + "')");
											try
											{ // close windows on desktop iParam
												CloseWindowsOnDesktop(iParam);
												rc = iParam;
											}
											catch
											{ // error while closing
												rc = -1;
											}
										}
										else
										{ // no desktop found
											if (verbose) Console.WriteLine("Could not find virtual desktop with name containing '" + groups[2].Value + "'");
											rc = -2;
										}
									}
								}
								break;

							case "GETDESKTOPFROMWINDOW": // get desktop from window
							case "GDFW":
								if (int.TryParse(groups[2].Value, out iParam))
								{ // check if parameter is an integer
									if (iParam > 0)
									{ // check if parameter is greater than 0
										try
										{ // seeking desktop for process handle
											iParam = (int)System.Diagnostics.Process.GetProcessById(iParam).MainWindowHandle;
											// process handle converted to window handle
											rc = VirtualDesktop.Desktop.FromDesktop(VirtualDesktop.Desktop.FromWindow((IntPtr)iParam));
											if (verbose) Console.WriteLine("Window is on desktop number " + rc.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(rc) + "')");
										}
										catch
										{ // error while seeking
											if (verbose) Console.WriteLine("Window to process id " + groups[2].Value + "' not found");
											rc = -1;
										}
									}
									else
										rc = -1;
								}
								else
								{
									try
									{ // seeking desktop for process name
										iParam = GetMainWindowHandle(groups[2].Value.Trim());
										rc = VirtualDesktop.Desktop.FromDesktop(VirtualDesktop.Desktop.FromWindow((IntPtr)iParam));
										if (verbose) Console.WriteLine("Window of process '" + groups[2].Value + "' is on desktop number " + rc.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(rc) + "')");
									}
									catch
									{ // error while seeking
										if (verbose) Console.WriteLine("Process '" + groups[2].Value + "' not found");
										rc = -1;
									}
								}
								break;

							case "GETDESKTOPFROMWINDOWHANDLE": // get desktop from window handle
							case "GDFWH":
								if (int.TryParse(groups[2].Value, out iParam))
								{ // check if parameter is an integer
									if (iParam > 0)
									{ // check if parameter is greater than 0
										try
										{ // seeking desktop for window handle
											rc = VirtualDesktop.Desktop.FromDesktop(VirtualDesktop.Desktop.FromWindow((IntPtr)iParam));
											if (verbose) Console.WriteLine("Window is on desktop number " + rc.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(rc) + "')");
										}
										catch
										{ // error while seeking
											if (verbose) Console.WriteLine("Window to handle " + groups[2].Value + "' not found");
											rc = -1;
										}
									}
									else
										rc = -1;
								}
								else
								{
									try
									{ // seeking window with window title
										iParam = (Int32)GetWindowFromTitle(groups[2].Value.Trim().Replace("^", ""));
										// seeking desktop for window handle
										rc = VirtualDesktop.Desktop.FromDesktop(VirtualDesktop.Desktop.FromWindow((IntPtr)iParam));
										if (verbose) Console.WriteLine("Window '" + foundTitle + "' is on desktop number " + rc.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(rc) + "')");
									}
									catch
									{ // error while seeking
										if (verbose) Console.WriteLine("Window with text '" + groups[2].Value + "' in title not found");
										rc = -1;
									}
								}
								break;

							case "ISWINDOWONDESKTOP": // is window on desktop in rc
							case "IWOD":
								if (int.TryParse(groups[2].Value, out iParam))
								{ // check if parameter is an integer
									if (iParam > 0)
									{ // check if parameter is greater than 0
										try
										{ // checking desktop for process handle
											iParam = (int)System.Diagnostics.Process.GetProcessById(iParam).MainWindowHandle;
											// process handle converted to window handle
											if (VirtualDesktop.Desktop.FromIndex(rc).HasWindow((IntPtr)iParam))
											{
												if (verbose) Console.WriteLine("Window to process id " + groups[2].Value + " is on desktop number " + rc.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(rc) + "')");
												rc = 0;
											}
											else
											{
												if (verbose) Console.WriteLine("Window to process id " + groups[2].Value + " is not on desktop number " + rc.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(rc) + "')");
												rc = 1;
											}
										}
										catch
										{ // error while seeking
											if (verbose) Console.WriteLine("Window to process id " + groups[2].Value + " not found");
											rc = -1;
										}
									}
									else
										rc = -1;
								}
								else
								{
									try
									{ // seeking desktop for process name
										iParam = GetMainWindowHandle(groups[2].Value.Trim());
										if (VirtualDesktop.Desktop.FromIndex(rc).HasWindow((IntPtr)iParam))
										{
											if (verbose) Console.WriteLine("Window of process '" + groups[2].Value + "' is on desktop number " + rc.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(rc) + "')");
											rc = 0;
										}
										else
										{
											if (verbose) Console.WriteLine("Window of process '" + groups[2].Value + "' is not on desktop number " + rc.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(rc) + "')");
											rc = 1;
										}
									}
									catch
									{ // error while seeking
										if (verbose) Console.WriteLine("Process '" + groups[2].Value + "' not found");
										rc = -1;
									}
								}
								break;

							case "ISWINDOWHANDLEONDESKTOP": // is window with handle on desktop in rc
							case "IWHOD":
								if (int.TryParse(groups[2].Value, out iParam))
								{ // check if parameter is an integer
									if (iParam > 0)
									{ // check if parameter is greater than 0
										try
										{ // checking desktop for window handle
											if (VirtualDesktop.Desktop.FromIndex(rc).HasWindow((IntPtr)iParam))
											{
												if (verbose) Console.WriteLine("Window to handle " + groups[2].Value + " is on desktop number " + rc.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(rc) + "')");
												rc = 0;
											}
											else
											{
												if (verbose) Console.WriteLine("Window to handle " + groups[2].Value + " is not on desktop number " + rc.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(rc) + "')");
												rc = 1;
											}
										}
										catch
										{ // error while seeking
											if (verbose) Console.WriteLine("Window to handle " + groups[2].Value + " not found");
											rc = -1;
										}
									}
									else
										rc = -1;
								}
								else
								{
									try
									{ // seeking window with window title
										iParam = (Int32)GetWindowFromTitle(groups[2].Value.Trim().Replace("^", ""));
										// checking desktop for window handle
										if (VirtualDesktop.Desktop.FromIndex(rc).HasWindow((IntPtr)iParam))
										{
											if (verbose) Console.WriteLine("Window '" + foundTitle + "' is on desktop number " + rc.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(rc) + "')");
											rc = 0;
										}
										else
										{
											if (verbose) Console.WriteLine("Window '" + foundTitle + "' is not on desktop number " + rc.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(rc) + "')");
											rc = 1;
										}
									}
									catch
									{ // error while seeking
										if (verbose) Console.WriteLine("Window with text '" + groups[2].Value + "' in title not found");
										rc = -1;
									}
								}
								break;

							case "MOVEWINDOW": // move window to desktop in rc
							case "MW":
								if (int.TryParse(groups[2].Value, out iParam))
								{ // check if parameter is an integer
									if (iParam > 0)
									{ // check if parameter is greater than 0
										try
										{ // seeking window for process handle
											iParam = (int)System.Diagnostics.Process.GetProcessById(iParam).MainWindowHandle;
											// process handle converted to window handle and move window
											VirtualDesktop.Desktop.FromIndex(rc).MoveWindow((IntPtr)iParam);
											if (verbose) Console.WriteLine("Window to process id " + groups[2].Value + " moved to desktop number " + rc.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(rc) + "')");
										}
										catch
										{ // error while seeking
											if (verbose) Console.WriteLine("Window to process id " + groups[2].Value + " not found or move failed");
											rc = -1;
										}
									}
									else
										rc = -1;
								}
								else
								{
									try
									{ // seeking window for process name
										iParam = GetMainWindowHandle(groups[2].Value.Trim());
										// move window
										VirtualDesktop.Desktop.FromIndex(rc).MoveWindow((IntPtr)iParam);
										if (verbose) Console.WriteLine("Window of process '" + groups[2].Value + "' moved to desktop number " + rc.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(rc) + "')");
									}
									catch
									{ // error while seeking
										if (verbose) Console.WriteLine("Process '" + groups[2].Value + "' not found or move failed");
										rc = -1;
									}
								}
								break;

							case "MOVEWINDOWHANDLE": // move window with handle to desktop in rc
							case "MWH":
								if (int.TryParse(groups[2].Value, out iParam))
								{ // check if parameter is an integer
									if (iParam > 0)
									{ // check if parameter is greater than 0
										try
										{
											// use window handle and move window
											VirtualDesktop.Desktop.FromIndex(rc).MoveWindow((IntPtr)iParam);
											if (verbose) Console.WriteLine("Window to handle " + groups[2].Value + " moved to desktop number " + rc.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(rc) + "')");
										}
										catch
										{ // error while seeking
											if (verbose) Console.WriteLine("Window to handle " + groups[2].Value + " not found or move failed");
											rc = -1;
										}
									}
									else
										rc = -1;
								}
								else
								{
									try
									{ // seeking window with window title
										iParam = (Int32)GetWindowFromTitle(groups[2].Value.Trim().Replace("^", ""));
										// move window
										VirtualDesktop.Desktop.FromIndex(rc).MoveWindow((IntPtr)iParam);
										if (verbose) Console.WriteLine("Window '" + foundTitle + "' moved to desktop number " + rc.ToString() + " (desktop '" + VirtualDesktop.Desktop.DesktopNameFromIndex(rc) + "')");
									}
									catch
									{ // error while seeking
										if (verbose) Console.WriteLine("Window with text '" + groups[2].Value + "' in title not found or move failed");
										rc = -1;
									}
								}
								break;

							case "ISWINDOWPINNED": // is window pinned to all desktops
							case "IWP":
								if (int.TryParse(groups[2].Value, out iParam))
								{ // check if parameter is an integer
									if (iParam > 0)
									{ // check if parameter is greater than 0
										try
										{ // checking desktop for process handle
											iParam = (int)System.Diagnostics.Process.GetProcessById(iParam).MainWindowHandle;
											// process handle converted to window handle
											if (VirtualDesktop.Desktop.IsWindowPinned((IntPtr)iParam))
											{
												if (verbose) Console.WriteLine("Window to process id " + groups[2].Value + " is pinned to all desktops");
												rc = 0;
											}
											else
											{
												if (verbose) Console.WriteLine("Window to process id " + groups[2].Value + " is not pinned to all desktops");
												rc = 1;
											}
										}
										catch
										{ // error while seeking
											if (verbose) Console.WriteLine("Window to process id " + groups[2].Value + " not found");
											rc = -1;
										}
									}
									else
										rc = -1;
								}
								else
								{
									try
									{ // seeking desktop for process name
										iParam = GetMainWindowHandle(groups[2].Value.Trim());
										if (VirtualDesktop.Desktop.IsWindowPinned((IntPtr)iParam))
										{
											if (verbose) Console.WriteLine("Window of process '" + groups[2].Value + "' is pinned to all desktops");
											rc = 0;
										}
										else
										{
											if (verbose) Console.WriteLine("Window of process '" + groups[2].Value + "' is not pinned to all desktops");
											rc = 1;
										}
									}
									catch
									{ // error while seeking
										if (verbose) Console.WriteLine("Process '" + groups[2].Value + "' not found");
										rc = -1;
									}
								}
								break;

							case "ISWINDOWHANDLEPINNED": // is window with handle pinned to all desktops
							case "IWHP":
								if (int.TryParse(groups[2].Value, out iParam))
								{ // check if parameter is an integer
									if (iParam > 0)
									{ // check if parameter is greater than 0
										try
										{ // checking desktop for window handle
											if (VirtualDesktop.Desktop.IsWindowPinned((IntPtr)iParam))
											{
												if (verbose) Console.WriteLine("Window to handle " + groups[2].Value + " is pinned to all desktops");
												rc = 0;
											}
											else
											{
												if (verbose) Console.WriteLine("Window to handle " + groups[2].Value + " is not pinned to all desktops");
												rc = 1;
											}
										}
										catch
										{ // error while seeking
											if (verbose) Console.WriteLine("Window to handle " + groups[2].Value + " not found");
											rc = -1;
										}
									}
									else
										rc = -1;
								}
								else
								{
									try
									{ // seeking window with window title
										iParam = (Int32)GetWindowFromTitle(groups[2].Value.Trim().Replace("^", ""));
										if (VirtualDesktop.Desktop.IsWindowPinned((IntPtr)iParam))
										{
											if (verbose) Console.WriteLine("Window '" + foundTitle + "' is pinned to all desktops");
											rc = 0;
										}
										else
										{
											if (verbose) Console.WriteLine("Window '" + foundTitle + "' is not pinned to all desktops");
											rc = 1;
										}
									}
									catch
									{ // error while seeking
										if (verbose) Console.WriteLine("Window with text '" + groups[2].Value + "' in title not found");
										rc = -1;
									}
								}
								break;

							case "PINWINDOW": // pin window to all desktops
							case "PW":
								if (int.TryParse(groups[2].Value, out iParam))
								{ // check if parameter is an integer
									if (iParam > 0)
									{ // check if parameter is greater than 0
										try
										{ // seeking window for process handle
											iParam = (int)System.Diagnostics.Process.GetProcessById(iParam).MainWindowHandle;
											// process handle converted to window handle and pin window
											VirtualDesktop.Desktop.PinWindow((IntPtr)iParam);
											if (verbose) Console.WriteLine("Window to process id " + groups[2].Value + " pinned to all desktops");
										}
										catch
										{ // error while seeking
											if (verbose) Console.WriteLine("Window to process id " + groups[2].Value + " not found or pin failed");
											rc = -1;
										}
									}
									else
										rc = -1;
								}
								else
								{
									try
									{ // seeking window for process name
										iParam = GetMainWindowHandle(groups[2].Value.Trim());
										// pin window
										VirtualDesktop.Desktop.PinWindow((IntPtr)iParam);
										if (verbose) Console.WriteLine("Window of process '" + groups[2].Value + "' pinned to all desktops");
									}
									catch
									{ // error while seeking
										if (verbose) Console.WriteLine("Process '" + groups[2].Value + "' not found or pin failed");
										rc = -1;
									}
								}
								break;

							case "PINWINDOWHANDLE": // pin window with handle to all desktops
							case "PWH":
								if (int.TryParse(groups[2].Value, out iParam))
								{ // check if parameter is an integer
									if (iParam > 0)
									{ // check if parameter is greater than 0
										try
										{ // process window handle and pin window
											VirtualDesktop.Desktop.PinWindow((IntPtr)iParam);
											if (verbose) Console.WriteLine("Window to handle " + groups[2].Value + " pinned to all desktops");
										}
										catch
										{ // error while seeking
											if (verbose) Console.WriteLine("Window to handle " + groups[2].Value + " not found or pin failed");
											rc = -1;
										}
									}
									else
										rc = -1;
								}
								else
								{
									try
									{ // seeking window with window title
										iParam = (Int32)GetWindowFromTitle(groups[2].Value.Trim().Replace("^", ""));
										// pin window
										VirtualDesktop.Desktop.PinWindow((IntPtr)iParam);
										if (verbose) Console.WriteLine("Window '" + foundTitle + "' pinned to all desktops");
									}
									catch
									{ // error while seeking
										if (verbose) Console.WriteLine("Window with text '" + groups[2].Value + "' in title not found or pin failed");
										rc = -1;
									}
								}
								break;

							case "UNPINWINDOW": // unpin window from all desktops
							case "UPW":
								if (int.TryParse(groups[2].Value, out iParam))
								{ // check if parameter is an integer
									if (iParam > 0)
									{ // check if parameter is greater than 0
										try
										{ // seeking window for process handle
											iParam = (int)System.Diagnostics.Process.GetProcessById(iParam).MainWindowHandle;
											// process handle converted to window handle and unpin window
											VirtualDesktop.Desktop.UnpinWindow((IntPtr)iParam);
											if (verbose) Console.WriteLine("Window to process id " + groups[2].Value + " unpinned from all desktops");
										}
										catch
										{ // error while seeking
											if (verbose) Console.WriteLine("Window to process id " + groups[2].Value + " not found or unpin failed");
											rc = -1;
										}
									}
									else
										rc = -1;
								}
								else
								{
									try
									{ // seeking window for process name
										iParam = GetMainWindowHandle(groups[2].Value.Trim());
										// unpin window
										VirtualDesktop.Desktop.UnpinWindow((IntPtr)iParam);
										if (verbose) Console.WriteLine("Window of process '" + groups[2].Value + "' unpinned from all desktops");
									}
									catch
									{ // error while seeking
										if (verbose) Console.WriteLine("Process '" + groups[2].Value + "' not found or unpin failed");
										rc = -1;
									}
								}
								break;

							case "UNPINWINDOWHANDLE": // unpin window with handle from all desktops
							case "UPWH":
								if (int.TryParse(groups[2].Value, out iParam))
								{ // check if parameter is an integer
									if (iParam > 0)
									{ // check if parameter is greater than 0
										try
										{ // process window handle and unpin window
											VirtualDesktop.Desktop.UnpinWindow((IntPtr)iParam);
											if (verbose) Console.WriteLine("Window to handle " + groups[2].Value + " unpinned from all desktops");
										}
										catch
										{ // error while seeking
											if (verbose) Console.WriteLine("Window to handle " + groups[2].Value + " not found or unpin failed");
											rc = -1;
										}
									}
									else
										rc = -1;
								}
								else
								{
									try
									{ // seeking window with window title
										iParam = (Int32)GetWindowFromTitle(groups[2].Value.Trim().Replace("^", ""));
										// unpin window
										VirtualDesktop.Desktop.UnpinWindow((IntPtr)iParam);
										if (verbose) Console.WriteLine("Window '" + foundTitle + "' unpinned from all desktops");
									}
									catch
									{ // error while seeking
										if (verbose) Console.WriteLine("Window with text '" + groups[2].Value + "' in title not found or unpin failed");
										rc = -1;
									}
								}
								break;

							case "ISAPPLICATIONPINNED": // is application pinned to all desktops
							case "IAP":
								if (int.TryParse(groups[2].Value, out iParam))
								{ // check if parameter is an integer
									if (iParam > 0)
									{ // check if parameter is greater than 0
										try
										{ // checking desktop for process handle
											iParam = (int)System.Diagnostics.Process.GetProcessById(iParam).MainWindowHandle;
											// process handle converted to window handle
											if (VirtualDesktop.Desktop.IsApplicationPinned((IntPtr)iParam))
											{
												if (verbose) Console.WriteLine("Application to process id " + groups[2].Value + " is pinned to all desktops");
												rc = 0;
											}
											else
											{
												if (verbose) Console.WriteLine("Application to process id " + groups[2].Value + " is not pinned to all desktops");
												rc = 1;
											}
										}
										catch
										{ // error while seeking
											if (verbose) Console.WriteLine("Window to process id " + groups[2].Value + " not found");
											rc = -1;
										}
									}
									else
										rc = -1;
								}
								else
								{
									try
									{ // seeking desktop for process name
										iParam = GetMainWindowHandle(groups[2].Value.Trim());
										if (VirtualDesktop.Desktop.IsApplicationPinned((IntPtr)iParam))
										{
											if (verbose) Console.WriteLine("Application of process '" + groups[2].Value + "' is pinned to all desktops");
											rc = 0;
										}
										else
										{
											if (verbose) Console.WriteLine("Application of process '" + groups[2].Value + "' is not pinned to all desktops");
											rc = 1;
										}
									}
									catch
									{ // error while seeking
										if (verbose) Console.WriteLine("Process '" + groups[2].Value + "' not found");
										rc = -1;
									}
								}
								break;

							case "PINAPPLICATION": // pin application to all desktops
							case "PA":
								if (int.TryParse(groups[2].Value, out iParam))
								{ // check if parameter is an integer
									if (iParam > 0)
									{ // check if parameter is greater than 0
										try
										{ // seeking window for process handle
											iParam = (int)System.Diagnostics.Process.GetProcessById(iParam).MainWindowHandle;
											// process handle converted to window handle and pin window
											VirtualDesktop.Desktop.PinApplication((IntPtr)iParam);
											if (verbose) Console.WriteLine("Application to process id " + groups[2].Value + " pinned to all desktops");
										}
										catch
										{ // error while seeking
											if (verbose) Console.WriteLine("Window to process id " + groups[2].Value + " not found or pin failed");
											rc = -1;
										}
									}
									else
										rc = -1;
								}
								else
								{
									try
									{ // seeking window for process name
										iParam = GetMainWindowHandle(groups[2].Value.Trim());
										// pin window
										VirtualDesktop.Desktop.PinApplication((IntPtr)iParam);
										if (verbose) Console.WriteLine("Application of process '" + groups[2].Value + "' pinned to all desktops");
									}
									catch
									{ // error while seeking
										if (verbose) Console.WriteLine("Process '" + groups[2].Value + "' not found or pin failed");
										rc = -1;
									}
								}
								break;

							case "UNPINAPPLICATION": // unpin application from all desktops
							case "UPA":
								if (int.TryParse(groups[2].Value, out iParam))
								{ // check if parameter is an integer
									if (iParam > 0)
									{ // check if parameter is greater than 0
										try
										{ // seeking window for process handle
											iParam = (int)System.Diagnostics.Process.GetProcessById(iParam).MainWindowHandle;
											// process handle converted to window handle and unpin window
											VirtualDesktop.Desktop.UnpinApplication((IntPtr)iParam);
											if (verbose) Console.WriteLine("Application to process id " + groups[2].Value + " unpinned from all desktops");
										}
										catch
										{ // error while seeking
											if (verbose) Console.WriteLine("Window to process id " + groups[2].Value + " not found or unpin failed");
											rc = -1;
										}
									}
									else
										rc = -1;
								}
								else
								{
									try
									{ // seeking window for process name
										iParam = GetMainWindowHandle(groups[2].Value.Trim());
										// unpin window
										VirtualDesktop.Desktop.UnpinApplication((IntPtr)iParam);
										if (verbose) Console.WriteLine("Application of process '" + groups[2].Value + "' unpinned from all desktops");
									}
									catch
									{ // error while seeking
										if (verbose) Console.WriteLine("Process '" + groups[2].Value + "' not found or unpin failed");
										rc = -1;
									}
								}
								break;

							case "CALCULATE": //calculate
							case "CALC":
							case "CA":
								if (int.TryParse(groups[2].Value, out iParam))
								{ // check if parameter is an integer
									if (verbose) Console.WriteLine("Adding " + iParam.ToString() + " to last result.");
									// adding iParam to result
									rc += iParam;
								}
								else
									rc = -2;
								break;

							case "SLEEP": //wait
							case "SL":
								if (int.TryParse(groups[2].Value, out iParam))
								{ // check if parameter is an integer
									if (iParam > 0)
									{ // check if parameter is greater than 0
										if (verbose) Console.WriteLine("Waiting " + iParam.ToString() + "ms");
										// waiting iParam milliseconds
										System.Threading.Thread.Sleep(iParam);
									}
									else
										rc = -1;
								}
								else
									rc = -2;
								break;

							default:
								rc = -2;
								break;
						}
					}
				}

				if (rc == -1)
				{ // error in action, stop processing
					Console.Error.WriteLine("Error while processing '" + arg + "'");
					if (breakonerror) break;
				}
				if (rc == -2)
				{ // error in parameter, stop processing
					Console.Error.WriteLine("Error in parameter '" + arg + "'");
					if (breakonerror) break;
				}
			}

			return rc;
		}

		static int GetMainWindowHandle(string ProcessName)
		{ // retrieve main window handle to process name
			System.Diagnostics.Process[] processes = System.Diagnostics.Process.GetProcessesByName(ProcessName);
			int wHwnd = 0;

			if (processes.Length > 0)
			{ // process found, get window handle
				wHwnd = (int)processes[0].MainWindowHandle;
			}

			return wHwnd;
		}

		private delegate bool EnumDelegate(IntPtr hWnd, int lParam);

		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = false)]
		static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, IntPtr lParam);

		static uint WM_CLOSE = 0x10;

		[DllImport("user32.dll")]
		[return: MarshalAs(UnmanagedType.Bool)]
		private static extern bool IsWindowVisible(IntPtr hWnd);

		[DllImport("user32.dll", CharSet=CharSet.Auto, SetLastError=true)]
		private static extern bool EnumDesktopWindows(IntPtr hDesktop, EnumDelegate lpEnumCallbackFunction, IntPtr lParam);

		[DllImport("user32.dll", CharSet=CharSet.Auto, SetLastError=true)]
		private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpWindowText, int nMaxCount);

		const int MAXTITLE = 255;
		private static IntPtr foundHandle;
		private static string foundTitle;
		private static string searchTitle;

		private static bool EnumWindowsProc(IntPtr hWnd, int lParam)
		{
			StringBuilder windowText = new StringBuilder(MAXTITLE);
			int titleLength = GetWindowText(hWnd, windowText, windowText.Capacity + 1);
			windowText.Length = titleLength;
			string title = windowText.ToString();

			if (!string.IsNullOrEmpty(title) && IsWindowVisible(hWnd))
			{
				if (title.ToUpper().IndexOf(searchTitle.ToUpper()) >= 0)
				{
					foundHandle = hWnd;
					foundTitle = title;
					return false;
				}
			}
			return true;
		}

		private static IntPtr GetWindowFromTitle(string searchFor)
		{
			searchTitle = searchFor;
			EnumDelegate enumfunc = new EnumDelegate(EnumWindowsProc);

			foundHandle = IntPtr.Zero;
			foundTitle = "";
			EnumDesktopWindows(IntPtr.Zero, enumfunc, IntPtr.Zero);
			if (foundHandle == IntPtr.Zero)
			{
				// Get the last Win32 error code
				int errorCode = Marshal.GetLastWin32Error();
				if (errorCode != 0)
				{ // error
					Console.WriteLine("EnumDesktopWindows failed with code {0}.", errorCode);
				}
			}
			return foundHandle;
		}

		private static int iListDesktop;

		private static bool EnumWindowsProcToList(IntPtr hWnd, int lParam)
		{
			try {
				int iDesktopIndex = VirtualDesktop.Desktop.FromDesktop(VirtualDesktop.Desktop.FromWindow(hWnd));
				if (iDesktopIndex == iListDesktop) Console.WriteLine(hWnd.ToInt32());
			}
			catch { }

			return true;
		}

		private static void ListWindowsOnDesktop(int DesktopIndex)
		{
			iListDesktop = DesktopIndex;
			EnumDelegate enumfunc = new EnumDelegate(EnumWindowsProcToList);

			EnumDesktopWindows(IntPtr.Zero, enumfunc, IntPtr.Zero);
		}

		private static int iCloseDesktop;

		private static bool EnumWindowsProcToClose(IntPtr hWnd, int lParam)
		{
			try {
				int iDesktopIndex = VirtualDesktop.Desktop.FromDesktop(VirtualDesktop.Desktop.FromWindow(hWnd));
				if (iDesktopIndex == iCloseDesktop) SendMessage(hWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
			}
			catch { }

			return true;
		}

		private static void CloseWindowsOnDesktop(int DesktopIndex)
		{
			iCloseDesktop = DesktopIndex;
			EnumDelegate enumfunc = new EnumDelegate(EnumWindowsProcToClose);

			EnumDesktopWindows(IntPtr.Zero, enumfunc, IntPtr.Zero);
		}

		private static int iMoveWindowsDesktop1;
		private static int iMoveWindowsDesktop2;

		private static bool EnumWindowsProcToMoveWindows(IntPtr hWnd, int lParam)
		{
			try {
				int iDesktopIndex = VirtualDesktop.Desktop.FromDesktop(VirtualDesktop.Desktop.FromWindow(hWnd));
				if (iDesktopIndex == iMoveWindowsDesktop1) VirtualDesktop.Desktop.FromIndex(iMoveWindowsDesktop2).MoveWindow(hWnd);
			}
			catch { }

			return true;
		}

		private static void MoveWindowsToDesktop(int MoveWindowsIndex1, int MoveWindowsIndex2)
		{
			iMoveWindowsDesktop1 = MoveWindowsIndex1;
			iMoveWindowsDesktop2 = MoveWindowsIndex2;
			EnumDelegate enumfunc = new EnumDelegate(EnumWindowsProcToMoveWindows);

			EnumDesktopWindows(IntPtr.Zero, enumfunc, IntPtr.Zero);
		}

		private static int iSwapDesktop1;
		private static int iSwapDesktop2;

		private static bool EnumWindowsProcToSwap(IntPtr hWnd, int lParam)
		{
			StringBuilder windowText = new StringBuilder(MAXTITLE);
			int titleLength = GetWindowText(hWnd, windowText, windowText.Capacity + 1);
			windowText.Length = titleLength;
			string title = windowText.ToString();

			if (!string.IsNullOrEmpty(title) && IsWindowVisible(hWnd))
			{
				try {
					int iDesktopIndex = VirtualDesktop.Desktop.FromDesktop(VirtualDesktop.Desktop.FromWindow(hWnd));
					if (iDesktopIndex == iSwapDesktop1) VirtualDesktop.Desktop.FromIndex(iSwapDesktop2).MoveWindow(hWnd);
					if (iDesktopIndex == iSwapDesktop2) VirtualDesktop.Desktop.FromIndex(iSwapDesktop1).MoveWindow(hWnd);
				}
				catch { }
			}

			return true;
		}

		private static void SwapDesktops(int SwapIndex1, int SwapIndex2)
		{
			iSwapDesktop1 = SwapIndex1;
			iSwapDesktop2 = SwapIndex2;
			EnumDelegate enumfunc = new EnumDelegate(EnumWindowsProcToSwap);

			EnumDesktopWindows(IntPtr.Zero, enumfunc, IntPtr.Zero);

			string desktopname1 = "";
			if (VirtualDesktop.Desktop.HasDesktopNameFromIndex(iSwapDesktop1))
				desktopname1 = VirtualDesktop.Desktop.DesktopNameFromIndex(iSwapDesktop1);
			string desktopname2 = "";
			if (VirtualDesktop.Desktop.HasDesktopNameFromIndex(iSwapDesktop2))
				desktopname2 = VirtualDesktop.Desktop.DesktopNameFromIndex(iSwapDesktop2);

			VirtualDesktop.Desktop.FromIndex(iSwapDesktop1).SetName(desktopname2);
			VirtualDesktop.Desktop.FromIndex(iSwapDesktop2).SetName(desktopname1);
		}

		private static int iInsertDesktop1;
		private static int iInsertDesktop2;

		private static bool EnumWindowsProcToInsert(IntPtr hWnd, int lParam)
		{
			StringBuilder windowText = new StringBuilder(MAXTITLE);
			int titleLength = GetWindowText(hWnd, windowText, windowText.Capacity + 1);
			windowText.Length = titleLength;
			string title = windowText.ToString();

			if (!string.IsNullOrEmpty(title) && IsWindowVisible(hWnd))
			{
				try {
					int iDesktopIndex = VirtualDesktop.Desktop.FromDesktop(VirtualDesktop.Desktop.FromWindow(hWnd));
					if ((iDesktopIndex >= iInsertDesktop1) && (iDesktopIndex < iInsertDesktop2))
						VirtualDesktop.Desktop.FromIndex(iDesktopIndex + 1).MoveWindow(hWnd);

					if (iDesktopIndex == iInsertDesktop2) VirtualDesktop.Desktop.FromIndex(iInsertDesktop1).MoveWindow(hWnd);
				}
				catch { }
			}

			return true;
		}

		private static void InsertDesktop(int InsertIndex1, int InsertIndex2)
		{
			if (InsertIndex2 > InsertIndex1)
			{
				iInsertDesktop1 = InsertIndex1;
				iInsertDesktop2 = InsertIndex2;
			}
			else
			{
				iInsertDesktop1 = InsertIndex2;
				iInsertDesktop2 = InsertIndex1;
			}
			EnumDelegate enumfunc = new EnumDelegate(EnumWindowsProcToInsert);

			EnumDesktopWindows(IntPtr.Zero, enumfunc, IntPtr.Zero);

			string desktopname1 = "";
			if (VirtualDesktop.Desktop.HasDesktopNameFromIndex(iInsertDesktop2))
				desktopname1 = VirtualDesktop.Desktop.DesktopNameFromIndex(iInsertDesktop2);

			for (int i = iInsertDesktop2 - 1; i >= iInsertDesktop1; i--)
			{
				string desktopname2 = "";
				if (VirtualDesktop.Desktop.HasDesktopNameFromIndex(i))
					desktopname2 = VirtualDesktop.Desktop.DesktopNameFromIndex(i);

				VirtualDesktop.Desktop.FromIndex(i + 1).SetName(desktopname2);
			}

			VirtualDesktop.Desktop.FromIndex(iInsertDesktop1).SetName(desktopname1);
		}

		static void HelpScreen()
		{
			Console.WriteLine("VirtualDesktop.exe\t\t\t\tMarkus Scholtes, 2023, v1.16\n");

			Console.WriteLine("Command line tool to manage the virtual desktops of Windows 10.");
			Console.WriteLine("Parameters can be given as a sequence of commands. The result - most of the");
			Console.WriteLine("times the number of the processed desktop - can be used as input for the next");
			Console.WriteLine("parameter. The result of the last command is returned as error level.");
			Console.WriteLine("Virtual desktop numbers start with 0.\n");
			Console.WriteLine("Parameters (leading / can be omitted or - can be used instead):\n");
			Console.WriteLine("/Help /h /?      this help screen.");
			Console.WriteLine("/Verbose /Quiet  enable verbose (default) or quiet mode (short: /v and /q).");
			Console.WriteLine("/Break /Continue break (default) or continue on error (short: /b and /co).");
			Console.WriteLine("/List            list all virtual desktops (short: /li).");
			Console.WriteLine("/Count           get count of virtual desktops to pipeline (short: /c).");
			Console.WriteLine("/GetDesktop:<n|s> get number of virtual desktop <n> or desktop with text <s> in");
			Console.WriteLine("                   name to pipeline (short: /gd).");
			Console.WriteLine("/GetCurrentDesktop  get number of current desktop to pipeline (short: /gcd).");
			Console.WriteLine("/Name[:<s>]      set name of desktop with number in pipeline (short: /na).");
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
			Console.WriteLine("/InsertDesktop:<n|s>  insert desktop number <n> or desktop with text <s> in");
			Console.WriteLine("                   name before desktop in pipeline or vice versa (short: /id).");
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
			Console.WriteLine("/PinWindowHandle:<s|n>  pin window with text <s> in title or handle <n> to all");
			Console.WriteLine("                   desktops (short: /pwh).");
			Console.WriteLine("/UnPinWindow:<s|n>  unpin process with name <s> or id <n> from all desktops");
			Console.WriteLine("                   (short: /upw).");
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
			Console.WriteLine("/WaitDesktopChange wait for desktop to change (short: /wdc).");
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

	}
}
