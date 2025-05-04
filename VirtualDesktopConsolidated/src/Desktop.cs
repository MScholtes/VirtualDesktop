using System;

namespace VirtualDesktop.Consolidated
{
    public class Desktop
    {
        public static int Count => DesktopManager.ApiFacade.GetDesktopCount();
        public static Desktop Current => new Desktop(DesktopManager.ApiFacade.GetCurrentDesktopIndex());

        public int Index { get; }
        public Desktop(int index) { Index = index; }

        public string Name
        {
            get => DesktopManager.ApiFacade.GetDesktopName(Index);
            set => DesktopManager.ApiFacade.SetDesktopName(Index, value);
        }

        public void MakeVisible() => DesktopManager.ApiFacade.SwitchDesktop(Index);
        public void Remove(Desktop fallback = null) => DesktopManager.ApiFacade.RemoveDesktop(Index, fallback?.Index ?? 0);
        public void SetWallpaper(string path)
        {
            if (!DesktopManager.ApiFacade.SupportsWallpaperSetting)
                throw new NotSupportedException("Wallpaper setting not supported on this Windows version.");
            DesktopManager.ApiFacade.SetDesktopWallpaper(Index, path);
        }

        public static string DesktopWallpaperFromIndex(int index)
        {
            if (!DesktopManager.ApiFacade.SupportsWallpaperSetting)
                return string.Empty;
            try
            {
                // Try to get wallpaper path via COM if available
                // This requires casting to the correct API facade type
                // We'll try Win11 and Win11_24H2, which support wallpaper
                var api = DesktopManager.ApiFacade;
                var type = api.GetType().Name;
                if (type.Contains("Win11_24H2"))
                {
                    var method = api.GetType().GetMethod("GetDesktopName"); // Just to check type
                    var desktopsField = api.GetType().GetField("_managerInternal", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    if (desktopsField != null)
                    {
                        dynamic managerInternal = desktopsField.GetValue(api);
                        object desktops;
                        managerInternal.GetDesktops(out desktops);
                        object objDesktop;
                        ((dynamic)desktops).GetAt(index, typeof(object).GUID, out objDesktop);
                        return ((dynamic)objDesktop).GetWallpaperPath();
                    }
                }
                else if (type.Contains("Win11"))
                {
                    var desktopsField = api.GetType().GetField("_managerInternal", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                    if (desktopsField != null)
                    {
                        dynamic managerInternal = desktopsField.GetValue(api);
                        object desktops;
                        managerInternal.GetDesktops(out desktops);
                        object objDesktop;
                        ((dynamic)desktops).GetAt(index, typeof(object).GUID, out objDesktop);
                        return ((dynamic)objDesktop).GetWallpaperPath();
                    }
                }
            }
            catch { }
            return string.Empty;
        }
        // Add other merged/forwarded methods as needed
    }
} 