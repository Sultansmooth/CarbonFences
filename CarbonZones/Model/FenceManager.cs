using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Xml.Serialization;
using CarbonZones.Win32;

namespace CarbonZones.Model
{
    public class FenceManager
    {
        public static FenceManager Instance { get; } = new FenceManager();

        private const string MetaFileName = "__fence_metadata.xml";

        private readonly string basePath;
        private readonly List<FenceWindow> fenceWindows = new List<FenceWindow>();
        private bool fencesVisible = true;

        public FenceManager()
        {
            basePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CarbonZones");
            EnsureDirectoryExists(basePath);
        }

        public void RegisterWindow(FenceWindow window)
        {
            fenceWindows.Add(window);
        }

        public void UnregisterWindow(FenceWindow window)
        {
            fenceWindows.Remove(window);
        }

        public void ToggleFences()
        {
            fencesVisible = !fencesVisible;
            foreach (var window in fenceWindows)
            {
                window.Visible = fencesVisible;
            }
            DesktopUtil.SetDesktopIconsVisible(!fencesVisible ? false : true);
        }

        public void LoadFences()
        {
            foreach (var dir in Directory.EnumerateDirectories(basePath))
            {
                var metaFile = Path.Combine(dir, MetaFileName);
                if (!File.Exists(metaFile)) continue;
                var serializer = new XmlSerializer(typeof(FenceInfo));
                var reader = new StreamReader(metaFile);
                var fence = serializer.Deserialize(reader) as FenceInfo;
                reader.Close();

                var window = new FenceWindow(fence);
                window.Show();
            }
        }

        public void CreateFence(string name)
        {
            CreateFence(name, new Rectangle(100, 250, 300, 300));
        }

        public void CreateFence(string name, Rectangle bounds)
        {
            var fenceInfo = new FenceInfo(Guid.NewGuid())
            {
                Name = name,
                PosX = bounds.X,
                PosY = bounds.Y,
                Width = bounds.Width,
                Height = bounds.Height
            };

            UpdateFence(fenceInfo);
            var window = new FenceWindow(fenceInfo);
            window.Show();
        }

        public void RemoveFence(FenceInfo info)
        {
            Directory.Delete(GetFolderPath(info), true);
        }

        public void UpdateFence(FenceInfo fenceInfo)
        {
            var path = GetFolderPath(fenceInfo);
            EnsureDirectoryExists(path);

            var metaFile = Path.Combine(path, MetaFileName);
            var serializer = new XmlSerializer(typeof(FenceInfo));
            var writer = new StreamWriter(metaFile);
            serializer.Serialize(writer, fenceInfo);
            writer.Close();
        }

        public void UnhideAllDesktopIcons()
        {
            try
            {
                foreach (var dir in Directory.EnumerateDirectories(basePath))
                {
                    var metaFile = Path.Combine(dir, MetaFileName);
                    if (!File.Exists(metaFile)) continue;
                    var serializer = new XmlSerializer(typeof(FenceInfo));
                    using var reader = new StreamReader(metaFile);
                    var fence = serializer.Deserialize(reader) as FenceInfo;
                    if (fence == null) continue;

                    var allFiles = new List<string>(fence.Files);
                    foreach (var tab in fence.Tabs)
                        allFiles.AddRange(tab.Files);

                    foreach (var filePath in allFiles)
                    {
                        try
                        {
                            if (!File.Exists(filePath) && !Directory.Exists(filePath)) continue;
                            var attrs = File.GetAttributes(filePath);
                            if (attrs.HasFlag(FileAttributes.Hidden))
                                File.SetAttributes(filePath, attrs & ~FileAttributes.Hidden);
                        }
                        catch { }
                    }
                }
            }
            catch { }

            DesktopUtil.SetDesktopIconsVisible(true);
        }

        private void EnsureDirectoryExists(string dir)
        {
            var di = new DirectoryInfo(dir);
            if (!di.Exists)
                di.Create();
        }

        private string GetFolderPath(FenceInfo fenceInfo)
        {
            return Path.Combine(basePath, fenceInfo.Id.ToString());
        }
    }
}
