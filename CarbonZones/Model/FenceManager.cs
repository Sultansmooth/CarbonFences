using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
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
        private System.Threading.Timer reconcileTimer;

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
                stagingWatcher.Renamed += OnStagingChanged;
                stagingWatcher.Created += OnStagingChanged;
                stagingWatcher.Deleted += OnStagingChanged;
            }
            catch { }
        }

        private void OnStagingChanged(object sender, FileSystemEventArgs e)
        {
            // Debounce: wait 500ms for operations to settle (delete+create = rename)
            reconcileTimer?.Dispose();
            reconcileTimer = new System.Threading.Timer(_ => ReconcileStaging(), null, 500, System.Threading.Timeout.Infinite);
        }

        private void ReconcileStaging()
        {
            try
            {
                if (!Directory.Exists(stagingDir)) return;

                // Build a set of all filenames currently in staging
                var stagedFiles = new HashSet<string>(
                    Directory.GetFileSystemEntries(stagingDir).Select(Path.GetFileName),
                    StringComparer.OrdinalIgnoreCase);

                // Build global tracked names once
                var allTrackedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var w in fenceWindows)
                    foreach (var t in w.FenceInfo.Tabs)
                        foreach (var f in t.Files)
                            allTrackedNames.Add(Path.GetFileName(f));

                foreach (var window in fenceWindows)
                {
                    bool changed = false;
                    foreach (var tab in window.FenceInfo.Tabs)
                    {
                        for (int i = tab.Files.Count - 1; i >= 0; i--)
                        {
                            var expectedStagedName = Path.GetFileName(tab.Files[i]);
                            if (stagedFiles.Contains(expectedStagedName))
                                continue; // file is still in staging under expected name

                            // Also check if it exists at original desktop path (not staged)
                            if (File.Exists(tab.Files[i]) || Directory.Exists(tab.Files[i]))
                                continue;

                            // File is missing from both staging and desktop — try to
                            // find an untracked file in staging (rename detection)
                            bool matched = false;
                            foreach (var stagedName in stagedFiles)
                            {
                                if (!allTrackedNames.Contains(stagedName))
                                {
                                    // Found an untracked file — this is the renamed version
                                    var dir = Path.GetDirectoryName(tab.Files[i]);
                                    allTrackedNames.Remove(expectedStagedName);
                                    tab.Files[i] = Path.Combine(dir, stagedName);
                                    allTrackedNames.Add(stagedName);
                                    changed = true;
                                    matched = true;
                                    break;
                                }
                            }

                            if (!matched)
                            {
                                // File was genuinely deleted — remove from fence
                                allTrackedNames.Remove(expectedStagedName);
                                tab.Files.RemoveAt(i);
                                changed = true;
                            }
                        }
                    }
                    if (changed)
                    {
                        UpdateFence(window.FenceInfo);
                        RefreshWindow(window);
                    }
                }
            }
            catch { }
        }

        private void RefreshWindow(FenceWindow window)
        {
            if (window.InvokeRequired)
                window.BeginInvoke(new Action(() => window.Refresh()));
            else
                window.Refresh();
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
