using System;
using Microsoft.Win32;

namespace VirtualDesktop.Consolidated
{
    public static class WindowsVersion
    {
        public enum WindowsApiVersion
        {
            Windows10_1607,
            Windows10_1809,
            Windows10_2004,
            Windows11_21H2,
            Windows11_22H2,
            Windows11_24H2,
            WindowsServer2016,
            WindowsServer2019,
            WindowsServer2022,
            Unknown
        }

        private static WindowsApiVersion? _apiVersion = null;

        public static WindowsApiVersion ApiVersion
        {
            get
            {
                if (_apiVersion == null)
                {
                    _apiVersion = DetectApiVersion();
                }
                return _apiVersion.Value;
            }
        }

        private static WindowsApiVersion DetectApiVersion()
        {
            try
            {
                using (var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion"))
                {
                    if (key != null)
                    {
                        var buildNumber = (string)key.GetValue("CurrentBuildNumber");
                        int build = 0;
                        int.TryParse(buildNumber, out build);

                        // Windows 11: build >= 22000
                        if (build >= 26100) return WindowsApiVersion.Windows11_24H2;
                        if (build >= 22621) return WindowsApiVersion.Windows11_22H2;
                        if (build >= 22000) return WindowsApiVersion.Windows11_21H2;
                        // Windows 10: build < 22000
                        if (build >= 19041) return WindowsApiVersion.Windows10_2004;
                        if (build >= 17763) return WindowsApiVersion.Windows10_1809;
                        if (build >= 14393) return WindowsApiVersion.Windows10_1607;
                        // Server
                        if (build >= 20348) return WindowsApiVersion.WindowsServer2022;
                        if (build >= 17763) return WindowsApiVersion.WindowsServer2019;
                        if (build >= 14393) return WindowsApiVersion.WindowsServer2016;
                    }
                }
            }
            catch { }
            return WindowsApiVersion.Unknown;
        }
    }
} 