using System;

namespace VirtualDesktop.Consolidated
{
    public static class DesktopManager
    {
        private static IVirtualDesktopApiFacade _apiFacade;
        public static IVirtualDesktopApiFacade ApiFacade
        {
            get
            {
                if (_apiFacade == null)
                {
                    _apiFacade = CreateApiFacade();
                }
                return _apiFacade;
            }
        }

        private static IVirtualDesktopApiFacade CreateApiFacade()
        {
            switch (WindowsVersion.ApiVersion)
            {
                case WindowsVersion.WindowsApiVersion.Windows10_1607:
                case WindowsVersion.WindowsApiVersion.Windows10_1809:
                case WindowsVersion.WindowsApiVersion.Windows10_2004:
                    return new VirtualDesktopApiFacade_Win10();
                case WindowsVersion.WindowsApiVersion.Windows11_21H2:
                case WindowsVersion.WindowsApiVersion.Windows11_22H2:
                    return new VirtualDesktopApiFacade_Win11();
                case WindowsVersion.WindowsApiVersion.Windows11_24H2:
                    return new VirtualDesktopApiFacade_Win11_24H2();
                case WindowsVersion.WindowsApiVersion.WindowsServer2016:
                case WindowsVersion.WindowsApiVersion.WindowsServer2019:
                case WindowsVersion.WindowsApiVersion.WindowsServer2022:
                    return new VirtualDesktopApiFacade_Server();
                default:
                    throw new NotSupportedException("Unsupported Windows version/build for VirtualDesktop.");
            }
        }
    }
} 