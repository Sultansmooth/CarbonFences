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
        private readonly string stagingDir;
        private readonly List<FenceWindow> fenceWindows = new List<FenceWindow>();
        private bool fencesVisible = true;
        private FileSystemWatcher stagingWatcher;

        public FenceManager()
        {
            basePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CarbonZones");
            stagingDir = Path.Combine(basePath, "__staged");
            EnsureDirectoryExists(basePath);
            StartStagingWatcher();
        }

        private void StartStagingWatcher()
        {
            try
            {
                EnsureDirectoryExists(stagingDir);
                stagingWatcher = new FileSystemWatcher(stagingDir)
                {
                    NotifyFilter = NotifyFilters.FileName | NotifyFilters.DirectoryName,
                    IncludeSubdirectories = false,
                    EnableRaisingEvents = true
                };
                stagingWatcher.Renamed += OnStagedFileRenamed;
            }
            catch { }
        }

        private void OnStagedFileRenamed(object sender, RenamedEventArgs e)
        {
            var oldFileName = Path.GetFileName(e.OldFullPath);
            var newFileName = Path.GetFileName(e.FullPath);

            foreach (var window in fenceWindows)
            {
                bool changed = false;
                foreach (var tab in window.FenceInfo.Tabs)
                {
                    for (int i = 0; i < tab.Files.Count; i++)
                    {
                        if (string.Equals(Path.GetFileName(tab.Files[i]), oldFileName, StringComparison.OrdinalIgnoreCase))
                        {
                            var dir = Path.GetDirectoryName(tab.Files[i]);
                            tab.Files[i] = Path.Combine(dir, newFileName);
                            changed = true;
                        }
                    }
                }
                if (changed)
                {
                    UpdateFence(window.FenceInfo);
                    if (window.InvokeRequired)
                        window.BeginInvoke(new Action(() => window.Refresh()));
                    else
                        window.Refresh();
                }
            }
        }

        public void RegisterWindow(FenceWindow window)
        {
            fenceWindows.Add(window);
        }

        public void UnregisterWindow(FenceWindow window)
        {
            fenceWindows.Remove(window);
        }

        public FenceWindow GetFenceAtScreenPoint(Point screenPoint)
        {
            foreach (var w in fenceWindows)
                if (w.Visible && w.Bounds.Contains(screenPoint))
                    return w;
            return null;
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
            // Unstage all files in this fence before deleting it
            var allFiles = new List<string>(info.Files);
            foreach (var tab in info.Tabs)
                allFiles.AddRange(tab.Files);
            foreach (var filePath in allFiles)
            {
                Win32.FileStaging.Unstage(filePath);
            }
            DesktopUtil.RefreshDesktopIcons();

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

        /// <summary>
        /// On startup, hide all files that are tracked by fences so they
        /// don't appear on the desktop alongside the fence windows.
        /// </summary>
        public void HideAllFencedIcons()
        {
            foreach (var filePath in EnumerateAllFencedFiles())
            {
                Win32.FileStaging.Stage(filePath);
            }
            DesktopUtil.RefreshDesktopIcons();
        }

        /// <summary>
        /// On exit, unhide all files tracked by fences so the desktop
        /// returns to its normal state.
        /// </summary>
        public void UnhideAllDesktopIcons()
        {
            foreach (var filePath in EnumerateAllFencedFiles())
            {
                Win32.FileStaging.Unstage(filePath);
            }
            DesktopUtil.RefreshDesktopIcons();
            DesktopUtil.SetDesktopIconsVisible(true);
        }

        private List<string> EnumerateAllFencedFiles()
        {
            var allFiles = new List<string>();
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

                    allFiles.AddRange(fence.Files);
                    foreach (var tab in fence.Tabs)
                        allFiles.AddRange(tab.Files);
                }
            }
            catch { }
            return allFiles;
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
