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
        // Add other merged/forwarded methods as needed
    }
} 