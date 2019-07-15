// Author: Markus Scholtes, 2019
// Version 1.4.1, 2019-07-15
// Version for Windows 10 1809
// Compile with:
// C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe VirtualDesktop.cs

using System;
using System.Runtime.InteropServices;
using System.ComponentModel;

// Based on http://stackoverflow.com/a/32417530, Windows 10 SDK and github project VirtualDesktop

namespace VirtualDesktop
{
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

	internal static class DesktopManager
	{
		static DesktopManager()
		{
			var shell = (IServiceProvider10)Activator.CreateInstance(Type.GetTypeFromCLSID(Guids.CLSID_ImmersiveShell));
			VirtualDesktopManagerInternal = (IVirtualDesktopManagerInternal)shell.QueryService(Guids.CLSID_VirtualDesktopManagerInternal, typeof(IVirtualDesktopManagerInternal).GUID);
			VirtualDesktopManager = (IVirtualDesktopManager)Activator.CreateInstance(Type.GetTypeFromCLSID(Guids.CLSID_VirtualDesktopManager));
			ApplicationViewCollection = (IApplicationViewCollection)shell.QueryService(typeof(IApplicationViewCollection).GUID, typeof(IApplicationViewCollection).GUID);
			VirtualDesktopPinnedApps = (IVirtualDesktopPinnedApps)shell.QueryService(Guids.CLSID_VirtualDesktopPinnedApps, typeof(IVirtualDesktopPinnedApps).GUID);
		}

		internal static IVirtualDesktopManagerInternal VirtualDesktopManagerInternal;
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


	public class Desktop
	{
		private IVirtualDesktop ivd;
		private Desktop(IVirtualDesktop desktop) { this.ivd = desktop; }

		public override int GetHashCode()
		{ // Get hash
			return ivd.GetHashCode();
		}

		public override bool Equals(object obj)
		{ // Compares with object
			var desk = obj as Desktop;
			return desk != null && object.ReferenceEquals(this.ivd, desk.ivd);
		}

		public static int Count
		{ // Returns the number of desktops
			get { return DesktopManager.VirtualDesktopManagerInternal.GetCount(); }
		}

		public static Desktop Current
		{ // Returns current desktop
			get { return new Desktop(DesktopManager.VirtualDesktopManagerInternal.GetCurrentDesktop()); }
		}

		public static Desktop FromIndex(int index)
		{ // Create desktop object from index 0..Count-1
			return new Desktop(DesktopManager.GetDesktop(index));
		}

		public static Desktop FromWindow(IntPtr hWnd)
		{ // Creates desktop object on which window <hWnd> is displayed
			if (hWnd == IntPtr.Zero) throw new ArgumentNullException();
			Guid id = DesktopManager.VirtualDesktopManager.GetWindowDesktopId(hWnd);
			return new Desktop(DesktopManager.VirtualDesktopManagerInternal.FindDesktop(ref id));
		}

		public static int FromDesktop(Desktop desktop)
		{ // Returns index of desktop object or -1 if not found
			return DesktopManager.GetDesktopIndex(desktop.ivd);
		}

		public static Desktop Create()
		{ // Create a new desktop
			return new Desktop(DesktopManager.VirtualDesktopManagerInternal.CreateDesktop());
		}

		public void Remove(Desktop fallback = null)
		{ // Destroy desktop and switch to <fallback>
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

		public bool IsVisible
		{ // Returns <true> if this desktop is the current displayed one
			get { return object.ReferenceEquals(ivd, DesktopManager.VirtualDesktopManagerInternal.GetCurrentDesktop()); }
		}

		public void MakeVisible()
		{ // Make this desktop visible
			DesktopManager.VirtualDesktopManagerInternal.SwitchDesktop(ivd);
		}

		public Desktop Left
		{ // Returns desktop at the left of this one, null if none
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
		{ // Returns desktop at the right of this one, null if none
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
		{ // Move window <hWnd> to this desktop
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
				DesktopManager.VirtualDesktopManagerInternal.MoveViewToDesktop(view, ivd);
			}
		}

		public void MoveActiveWindow()
		{
			MoveWindow(GetForegroundWindow());
		}

		public bool HasWindow(IntPtr hWnd)
		{ // Returns true if window <hWnd> is on this desktop
			if (hWnd == IntPtr.Zero) throw new ArgumentNullException();
			return ivd.GetId() == DesktopManager.VirtualDesktopManager.GetWindowDesktopId(hWnd);
		}

		public static bool IsWindowPinned(IntPtr hWnd)
		{ // Returns true if window <hWnd> is pinned to all desktops
			if (hWnd == IntPtr.Zero) throw new ArgumentNullException();
			return DesktopManager.VirtualDesktopPinnedApps.IsViewPinned(hWnd.GetApplicationView());
		}

		public static void PinWindow(IntPtr hWnd)
		{ // pin window <hWnd> to all desktops
			if (hWnd == IntPtr.Zero) throw new ArgumentNullException();
			var view = hWnd.GetApplicationView();
			if (!DesktopManager.VirtualDesktopPinnedApps.IsViewPinned(view))
			{ // pin only if not already pinned
				DesktopManager.VirtualDesktopPinnedApps.PinView(view);
			}
		}

		public static void UnpinWindow(IntPtr hWnd)
		{ // unpin window <hWnd> from all desktops
			if (hWnd == IntPtr.Zero) throw new ArgumentNullException();
			var view = hWnd.GetApplicationView();
			if (DesktopManager.VirtualDesktopPinnedApps.IsViewPinned(view))
			{ // unpin only if not already unpinned
				DesktopManager.VirtualDesktopPinnedApps.UnpinView(view);
			}
		}

		public static bool IsApplicationPinned(IntPtr hWnd)
		{ // Returns true if application for window <hWnd> is pinned to all desktops
			if (hWnd == IntPtr.Zero) throw new ArgumentNullException();
			return DesktopManager.VirtualDesktopPinnedApps.IsAppIdPinned(DesktopManager.GetAppId(hWnd));
		}

		public static void PinApplication(IntPtr hWnd)
		{ // pin application for window <hWnd> to all desktops
			if (hWnd == IntPtr.Zero) throw new ArgumentNullException();
			string appId = DesktopManager.GetAppId(hWnd);
			if (!DesktopManager.VirtualDesktopPinnedApps.IsAppIdPinned(appId))
			{ // pin only if not already pinned
				DesktopManager.VirtualDesktopPinnedApps.PinAppID(appId);
			}
		}

		public static void UnpinApplication(IntPtr hWnd)
		{ // unpin application for window <hWnd> from all desktops
			if (hWnd == IntPtr.Zero) throw new ArgumentNullException();
			var view = hWnd.GetApplicationView();
			string appId = DesktopManager.GetAppId(hWnd);
			if (DesktopManager.VirtualDesktopPinnedApps.IsAppIdPinned(appId))
			{ // unpin only if already pinned
				DesktopManager.VirtualDesktopPinnedApps.UnpinAppID(appId);
			}
		}

		// Get process id to window handle
		[DllImport("user32.dll")]
		public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

		// Get handle of active window
		[DllImport("user32.dll")]
		public static extern IntPtr GetForegroundWindow();
	}
}


namespace VDeskTool
{
  static class Program
  {
  	static bool verbose = true;
  	static bool breakonerror = true;
  	static int rc = 0;

    static int Main(string[] args)
    {
      if (args.Length == 0)
      { // no arguments, show help screen
      	HelpScreen();
      	return -2;
      }

			foreach (string arg in args)
			{
     		System.Text.RegularExpressions.GroupCollection groups = System.Text.RegularExpressions.Regex.Match(arg, @"^[-\/]?([^:=]+)[:=]?([^:=]*)$").Groups;

				if (groups.Count != 3)
				{ // parameter error
					rc = -2;
				}
				else
				{ // reset return code if on error
					if (rc < 0) rc = 0;

					if (groups[2].Value == "")
					{ // parameter without value
						switch(groups[1].Value.ToUpper())
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

							case "COUNT": // get count of desktops
							case "C":
								rc = VirtualDesktop.Desktop.Count;
								if (verbose) Console.WriteLine("Count of desktops: " + rc);
								break;

							case "GETCURRENTDESKTOP": // get number of current desktop
							case "GCD":
								rc = VirtualDesktop.Desktop.FromDesktop(VirtualDesktop.Desktop.Current);
								if (verbose) Console.WriteLine("Number of current desktop: " + rc);
								break;

							case "ISVISIBLE": // is desktop in rc visible?
							case "IV":
								if ((rc >= 0) && (rc < VirtualDesktop.Desktop.Count))
								{ // check if parameter is 0 and in range of active desktops
									if (VirtualDesktop.Desktop.FromIndex(rc).IsVisible)
									{
										if (verbose) Console.WriteLine("Virtual desktop " + rc.ToString() + " is visible");
										rc = 0;
									}
									else
									{
										if (verbose) Console.WriteLine("Virtual desktop " + rc.ToString() + " is not visible");
										rc = 1;
									}
								}
								else
									rc = -1;
								break;

							case "SWITCH": // switch to desktop in rc
							case "S":
								if (verbose) Console.WriteLine("Switching to virtual desktop " + rc.ToString());
								try
								{ // activate virtual desktop rc
									VirtualDesktop.Desktop.FromIndex(rc).MakeVisible();
								}
								catch
								{ // error while activating
									rc = -1;
								}
								break;

							case "LEFT": // switch to desktop to the left
							case "L":
								if (verbose) Console.Write("Switching to left virtual desktop");
								try
								{ // activate virtual desktop to the left
									VirtualDesktop.Desktop.Current.Left.MakeVisible();
									rc = VirtualDesktop.Desktop.FromDesktop(VirtualDesktop.Desktop.Current);
									if (verbose) Console.WriteLine(", desktop " + rc.ToString() + " is active now");
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
									VirtualDesktop.Desktop.Current.Right.MakeVisible();
									rc = VirtualDesktop.Desktop.FromDesktop(VirtualDesktop.Desktop.Current);
									if (verbose) Console.WriteLine(", desktop " + rc.ToString() + " is active now");
								}
								catch
								{ // error while activating
									if (verbose) Console.WriteLine();
									rc = -1;
								}
								break;

							case "NEW": // create new desktop
							case "N":
								if (verbose) Console.Write("Creating virtual desktop ");
								try
								{ // create virtual desktop, number is stored in rc
									rc = VirtualDesktop.Desktop.FromDesktop(VirtualDesktop.Desktop.Create());
									if (verbose) Console.WriteLine(rc.ToString());
								}
								catch
								{ // error while creating
									Console.WriteLine();
									rc = -1;
								}
								break;

							case "REMOVE": // remove desktop in rc
							case "R":
								if (verbose) Console.WriteLine("Removing virtual desktop " + rc.ToString());
								try
								{ // remove virtual desktop rc
									VirtualDesktop.Desktop.FromIndex(rc).Remove();
								}
								catch
								{ // error while removing
									rc = -1;
								}
								break;

							case "MOVEACTIVEWINDOW": // move active window to desktop in rc
							case "MAW":
								if (verbose) Console.WriteLine("Moving active window to virtual desktop " + rc.ToString());
								try
								{ // move active window
									VirtualDesktop.Desktop.FromIndex(rc).MoveActiveWindow();
									if (verbose) Console.WriteLine("Active window moved to desktop " + rc.ToString());
								}
								catch
								{ // error while moving
									if (verbose) Console.WriteLine("No active window or move failed");
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
								{ // check if parameter is an integer
									if ((iParam >= 0) && (iParam < VirtualDesktop.Desktop.Count))
									{ // check if parameter is 0 and in range of active desktops
										if (verbose) Console.WriteLine("Virtual desktop " + iParam.ToString() + " selected");
										rc = iParam;
									}
									else
										rc = -1;
								}
								else
									rc = -2;
								break;

							case "ISVISIBLE": // is desktop visible?
							case "IV":
								if (int.TryParse(groups[2].Value, out iParam))
								{ // check if parameter is an integer
									if ((iParam >= 0) && (iParam < VirtualDesktop.Desktop.Count))
									{ // check if parameter is 0 and in range of active desktops
										if (VirtualDesktop.Desktop.FromIndex(iParam).IsVisible)
										{
											if (verbose) Console.WriteLine("Virtual desktop " + iParam.ToString() + " is visible");
											rc = 0;
										}
										else
										{
											if (verbose) Console.WriteLine("Virtual desktop " + iParam.ToString() + " is not visible");
											rc = 1;
										}
									}
									else
										rc = -1;
								}
								else
									rc = -2;
								break;

							case "SWITCH": // switch to desktop
							case "S":
								if (int.TryParse(groups[2].Value, out iParam))
								{ // check if parameter is an integer
									if ((iParam >= 0) && (iParam < VirtualDesktop.Desktop.Count))
									{ // check if parameter is 0 and in range of active desktops
										if (verbose) Console.WriteLine("Switching to virtual desktop " + iParam.ToString());
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
									rc = -2;
								break;

							case "REMOVE": // remove desktop
							case "R":
								if (int.TryParse(groups[2].Value, out iParam))
								{ // check if parameter is an integer
									if ((iParam >= 0) && (iParam < VirtualDesktop.Desktop.Count))
									{ // check if parameter is 0 and in range of active desktops
										if (verbose) Console.WriteLine("Removing virtual desktop " + iParam.ToString());
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
									rc = -2;
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
											if (verbose) Console.WriteLine("Window is on desktop " + rc.ToString());
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
										iParam = GetWindowHandle(groups[2].Value.Trim());
										rc = VirtualDesktop.Desktop.FromDesktop(VirtualDesktop.Desktop.FromWindow((IntPtr)iParam));
										if (verbose) Console.WriteLine("Window of process '" + groups[2].Value + "' is on desktop " + rc.ToString());
									}
									catch
									{ // error while seeking
										if (verbose) Console.WriteLine("Process '" + groups[2].Value + "' not found");
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
												if (verbose) Console.WriteLine("Window to process id " + groups[2].Value + " is on desktop " + rc.ToString());
												rc = 0;
											}
											else
											{
												if (verbose) Console.WriteLine("Window to process id " + groups[2].Value + " is not on desktop " + rc.ToString());
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
										iParam = GetWindowHandle(groups[2].Value.Trim());
										if (VirtualDesktop.Desktop.FromIndex(rc).HasWindow((IntPtr)iParam))
										{
											if (verbose) Console.WriteLine("Window of process '" + groups[2].Value + "' is on desktop " + rc.ToString());
											rc = 0;
										}
										else
										{
											if (verbose) Console.WriteLine("Window of process '" + groups[2].Value + "' is not on desktop " + rc.ToString());
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
											if (verbose) Console.WriteLine("Window to process id " + groups[2].Value + " moved to desktop " + rc.ToString());
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
										iParam = GetWindowHandle(groups[2].Value.Trim());
										// move window
										VirtualDesktop.Desktop.FromIndex(rc).MoveWindow((IntPtr)iParam);
										if (verbose) Console.WriteLine("Window of process '" + groups[2].Value + "' moved to desktop " + rc.ToString());
									}
									catch
									{ // error while seeking
										if (verbose) Console.WriteLine("Process '" + groups[2].Value + "' not found or move failed");
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
										iParam = GetWindowHandle(groups[2].Value.Trim());
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
										iParam = GetWindowHandle(groups[2].Value.Trim());
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
										iParam = GetWindowHandle(groups[2].Value.Trim());
										// pin window
										VirtualDesktop.Desktop.UnpinWindow((IntPtr)iParam);
										if (verbose) Console.WriteLine("Window of process '" + groups[2].Value + "' unpinned from all desktops");
									}
									catch
									{ // error while seeking
										if (verbose) Console.WriteLine("Process '" + groups[2].Value + "' not found or pin failed");
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
										iParam = GetWindowHandle(groups[2].Value.Trim());
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
										iParam = GetWindowHandle(groups[2].Value.Trim());
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
										iParam = GetWindowHandle(groups[2].Value.Trim());
										// pin window
										VirtualDesktop.Desktop.UnpinApplication((IntPtr)iParam);
										if (verbose) Console.WriteLine("Application of process '" + groups[2].Value + "' unpinned from all desktops");
									}
									catch
									{ // error while seeking
										if (verbose) Console.WriteLine("Process '" + groups[2].Value + "' not found or pin failed");
										rc = -1;
									}
								}
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

		static int GetWindowHandle(string ProcessName)
		{ // retrieve window handle to process name
			System.Diagnostics.Process[] processes = System.Diagnostics.Process.GetProcessesByName(ProcessName);
			int wHwnd = 0;

			if (processes.Length > 0)
			{ // process found, get window handle
				wHwnd = (int)processes[0].MainWindowHandle;
			}

			return wHwnd;
		}

    static void HelpScreen()
    {
    	Console.WriteLine("VirtualDesktop.exe\t\t\t\tMarkus Scholtes, 2019, v1.4.1\n");

    	Console.WriteLine("Command line tool to manage the virtual desktops of Windows 10.");
    	Console.WriteLine("Parameters can be given as a sequence of commands. The result - most of the");
    	Console.WriteLine("times the number of the processed desktop - can be used as input for the next");
    	Console.WriteLine("parameter. The result of the last command is returned as error level.");
    	Console.WriteLine("Virtual desktop numbers start with 0.\n");
    	Console.WriteLine("Parameters (leading / can be omitted or - can be used instead):\n");
    	Console.WriteLine("/Help /h /?      this help screen.");
    	Console.WriteLine("/Verbose /Quiet  enable verbose (default) or quiet mode (short: /v and /q).");
    	Console.WriteLine("/Break /Continue break (default) or continue on error (short: /b and /co).");
    	Console.WriteLine("/Count           get count of virtual desktops to pipeline (short: /c).");
    	Console.WriteLine("/GetDesktop:<n>  get number of virtual desktop <n> to pipeline (short: /gd).");
    	Console.WriteLine("/GetCurrentDesktop  get number of current desktop to pipeline (short: /gcd).");
    	Console.WriteLine("/IsVisible[:<n>]  is desktop number <n> or number in pipeline visible (short:");
    	Console.WriteLine("                    /iv)? Returns 0 for visible and 1 for invisible.");
    	Console.WriteLine("/Switch[:<n>]    switch to desktop with number <n> or with number in pipeline");
    	Console.WriteLine("                   (short: /s).");
    	Console.WriteLine("/Left            switch to virtual desktop to the left of the active desktop");
    	Console.WriteLine("                   (short: /l).");
    	Console.WriteLine("/Right           switch to virtual desktop to the right of the active desktop");
    	Console.WriteLine("                   (short: /ri).");
    	Console.WriteLine("/New             create new desktop (short: /n). Number is stored in pipeline.");
    	Console.WriteLine("/Remove[:<n>]    remove desktop number <n> or desktop with number in pipeline");
    	Console.WriteLine("                   (short: /r).");
    	Console.WriteLine("/MoveWindow:<s>  move process with name <s> to desktop with number in pipeline");
    	Console.WriteLine("                   (short: /mw).");
    	Console.WriteLine("/MoveWindow:<n>  move process with id <n> to desktop with number in pipeline");
    	Console.WriteLine("                   (short: /mw).");
    	Console.WriteLine("/MoveActiveWindow  move active window to desktop with number in pipeline");
    	Console.WriteLine("                   (short: /maw).");
    	Console.WriteLine("/GetDesktopFromWindow:<s>  get desktop number where process with name <s> is");
    	Console.WriteLine("                   displayed (short: /gdfw).");
    	Console.WriteLine("/GetDesktopFromWindow:<n>  get desktop number where process with id <n> is");
    	Console.WriteLine("                   displayed (short: /gdfw).");
    	Console.WriteLine("/IsWindowOnDesktop:<s>  check if process with name <s> is on desktop with num-");
    	Console.WriteLine("                   ber in pipeline (short: /iwod). Returns 0 for yes, 1 for no.");
    	Console.WriteLine("/IsWindowOnDesktop:<n>  check if process with id <n> is on desktop with number");
    	Console.WriteLine("                   in pipeline (short: /iwod). Returns 0 for yes, 1 for no.");
    	Console.WriteLine("/PinWindow:<s>   pin process with name <s> to all desktops (short: /pw).");
    	Console.WriteLine("/PinWindow:<n>   pin process with id <n> to all desktops (short: /pw).");
    	Console.WriteLine("/UnPinWindow:<s>  unpin process with name <s> from all desktops (short: /upw).");
    	Console.WriteLine("/UnPinWindow:<n>  unpin process with id <n> from all desktops (short: /upw).");
    	Console.WriteLine("/IsWindowPinned:<s>  check if process with name <s> is pinned to all desktops");
    	Console.WriteLine("                   (short: /iwp). Returns 0 for yes, 1 for no.");
    	Console.WriteLine("/IsWindowPinned:<n>  check if process with id <n> is pinned to all desktops");
    	Console.WriteLine("                   (short: /iwp). Returns 0 for yes, 1 for no.");
    	Console.WriteLine("/PinApplication:<s>  pin application with name <s> to all desktops (short: /pa)");
    	Console.WriteLine("/PinApplication:<n>  pin application with process id <n> to all desktops");
    	Console.WriteLine("                   (short: /pa).");
    	Console.WriteLine("/UnPinApplication:<s>  unpin application with name <s> from all desktops");
    	Console.WriteLine("                   (short: /upa).");
    	Console.WriteLine("/UnPinApplication:<n>  unpin application with process id <n> from all desktops");
    	Console.WriteLine("                   (short: /upa).");
    	Console.WriteLine("/IsApplicationPinned:<s>  check if application with name <s> is pinned to all");
    	Console.WriteLine("                   desktops (short: /iap). Returns 0 for yes, 1 for no.");
    	Console.WriteLine("/IsApplicationPinned:<n>  check if application with process id <n> is pinned to");
    	Console.WriteLine("                   all desktops (short: /iap). Returns 0 for yes, 1 for no.");
    	Console.WriteLine("/WaitKey         wait for key press (short: /wk).");
    	Console.WriteLine("/Sleep:<n>       wait for <n> milliseconds (short: /sl).\n");
    	Console.WriteLine("Examples:\n");
    	Console.WriteLine("Virtualdesktop.exe -New -Switch -GetCurrentDesktop");
    	Console.WriteLine("Virtualdesktop.exe sleep:200 gd:1 mw:notepad s");
    	Console.WriteLine("Virtualdesktop.exe /Count /continue /Remove /Remove /Count");
    	Console.WriteLine("VirtualDesktop.exe -IsWindowPinned:cmd");
    	Console.WriteLine("if ERRORLEVEL 1 VirtualDesktop.exe PinWindow:cmd");
    }

  }
}
