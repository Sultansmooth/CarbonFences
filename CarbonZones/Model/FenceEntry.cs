using System.Drawing;
using System.Threading.Tasks;
using System.Diagnostics;
using System;
using System.IO;
using CarbonZones.Win32;
using CarbonZones.Util;

namespace CarbonZones.Model
{
    public class FenceEntry
    {
        public string Path { get; }

        public EntryType Type { get; }

        public string Name => System.IO.Path.GetFileNameWithoutExtension(Path);

        private FenceEntry(string path, EntryType type)
        {
            Path = path;
            Type = type;
        }

        public static FenceEntry FromPath(string path)
        {
            if (File.Exists(path))
                return new FenceEntry(path, EntryType.File);
            else if (Directory.Exists(path))
                return new FenceEntry(path, EntryType.Folder);
            else return null;
        }

        public Icon ExtractIcon(ThumbnailProvider thumbnailProvider)
        {
            if (Type == EntryType.File)
            {
                if (thumbnailProvider.IsSupported(Path))
                    return thumbnailProvider.GenerateThumbnail(Path);
                else
                    return Icon.ExtractAssociatedIcon(Path);
            }
            else
            {
                return IconUtil.FolderLarge;
            }
        }

        public Bitmap ExtractLargeIcon(ThumbnailProvider thumbnailProvider)
        {
            if (Type == EntryType.File)
            {
                if (thumbnailProvider.IsSupported(Path))
                {
                    var icon = thumbnailProvider.GenerateThumbnail(Path);
                    return icon?.ToBitmap();
                }
                return IconUtil.GetJumboIcon(Path);
            }
            else
            {
                return IconUtil.GetFolderJumbo();
            }
        }

        public void Open()
        {
            Task.Run(() =>
            {
                // start asynchronously
                try
                {
                    if (Type == EntryType.File)
                        Process.Start(new ProcessStartInfo(Path) { UseShellExecute = true });
                    else if (Type == EntryType.Folder)
                        Process.Start(new ProcessStartInfo("explorer.exe", Path) { UseShellExecute = true });
                }
                catch (Exception e)
                {
                    Console.WriteLine($"Failed to start: {e}");
                }
            });
        }
    }
}
