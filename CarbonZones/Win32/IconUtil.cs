using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace CarbonZones.Win32
{
    public static class IconUtil
    {
        private static Icon folderIcon;
        private static Bitmap folderJumbo;

        public static Icon FolderLarge => folderIcon ?? (folderIcon = GetStockIcon(SHSIID_FOLDER, SHGSI_LARGEICON));

        public static Bitmap GetJumboIcon(string path)
        {
            try
            {
                var shfi = new SHFILEINFO();
                var result = SHGetFileInfo(path, 0, ref shfi,
                    (uint)Marshal.SizeOf(typeof(SHFILEINFO)),
                    SHGFI_SYSICONINDEX);
                if (result == IntPtr.Zero) return null;

                var iidImageList = IID_IImageList;
                int hr = SHGetImageList(SHIL_JUMBO, ref iidImageList, out IntPtr imgList);
                if (hr != 0 || imgList == IntPtr.Zero) return null;

                IntPtr hIcon = ImageList_GetIcon(imgList, shfi.iIcon, 0);
                if (hIcon == IntPtr.Zero) return null;

                using var icon = Icon.FromHandle(hIcon);
                var bmp = icon.ToBitmap();
                DestroyIcon(hIcon);
                return bmp;
            }
            catch
            {
                return null;
            }
        }

        public static Bitmap GetFolderJumbo()
        {
            if (folderJumbo != null) return folderJumbo;
            try
            {
                var shfi = new SHFILEINFO();
                var result = SHGetFileInfo("folder", FILE_ATTRIBUTE_DIRECTORY, ref shfi,
                    (uint)Marshal.SizeOf(typeof(SHFILEINFO)),
                    SHGFI_SYSICONINDEX | SHGFI_USEFILEATTRIBUTES);
                if (result == IntPtr.Zero) return null;

                var iidImageList = IID_IImageList;
                int hr = SHGetImageList(SHIL_JUMBO, ref iidImageList, out IntPtr imgList);
                if (hr != 0 || imgList == IntPtr.Zero) return null;

                IntPtr hIcon = ImageList_GetIcon(imgList, shfi.iIcon, 0);
                if (hIcon == IntPtr.Zero) return null;

                using var icon = Icon.FromHandle(hIcon);
                folderJumbo = icon.ToBitmap();
                DestroyIcon(hIcon);
                return folderJumbo;
            }
            catch
            {
                return null;
            }
        }

        private static Icon GetStockIcon(uint type, uint size)
        {
            var info = new SHSTOCKICONINFO();
            info.cbSize = (uint)Marshal.SizeOf(info);

            SHGetStockIconInfo(type, SHGSI_ICON | size, ref info);

            var icon = (Icon)Icon.FromHandle(info.hIcon).Clone();
            DestroyIcon(info.hIcon);

            return icon;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct SHSTOCKICONINFO
        {
            public uint cbSize;
            public IntPtr hIcon;
            public int iSysIconIndex;
            public int iIcon;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string szPath;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct SHFILEINFO
        {
            public IntPtr hIcon;
            public int iIcon;
            public uint dwAttributes;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string szDisplayName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
            public string szTypeName;
        }

        [DllImport("shell32.dll")]
        public static extern int SHGetStockIconInfo(uint siid, uint uFlags, ref SHSTOCKICONINFO psii);

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SHGetFileInfo(string pszPath, uint dwFileAttributes,
            ref SHFILEINFO psfi, uint cbSizeFileInfo, uint uFlags);

        [DllImport("shell32.dll")]
        private static extern int SHGetImageList(int iImageList, ref Guid riid, out IntPtr ppv);

        [DllImport("comctl32.dll")]
        private static extern IntPtr ImageList_GetIcon(IntPtr himl, int i, uint flags);

        [DllImport("user32.dll")]
        public static extern bool DestroyIcon(IntPtr handle);

        private const uint SHSIID_FOLDER = 0x3;
        private const uint SHGSI_ICON = 0x100;
        private const uint SHGSI_LARGEICON = 0x0;
        private const uint SHGSI_SMALLICON = 0x1;
        private const uint SHGFI_SYSICONINDEX = 0x4000;
        private const uint SHGFI_USEFILEATTRIBUTES = 0x10;
        private const uint FILE_ATTRIBUTE_DIRECTORY = 0x10;
        private const int SHIL_JUMBO = 0x4; // 256x256
        private static readonly Guid IID_IImageList = new Guid("46EB5926-582E-4017-9FDF-E8998DAA0950");
    }
}
