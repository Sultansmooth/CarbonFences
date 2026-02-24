using System;
using System.IO;

namespace CarbonZones.Win32
{
    public static class FileStaging
    {
        private static readonly string StagingDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "CarbonZones", "__staged");

        public static string GetStagedPath(string originalPath)
        {
            return Path.Combine(StagingDir, Path.GetFileName(originalPath));
        }

        /// <summary>
        /// Returns the actual filesystem path for a fenced file.
        /// If the file was staged (moved), returns the staged path.
        /// Otherwise returns the original path.
        /// </summary>
        public static string GetEffectivePath(string originalPath)
        {
            if (File.Exists(originalPath) || Directory.Exists(originalPath))
                return originalPath;
            var staged = GetStagedPath(originalPath);
            if (File.Exists(staged) || Directory.Exists(staged))
                return staged;
            return originalPath;
        }

        /// <summary>
        /// Move a file/folder from the desktop into the staging directory,
        /// making it invisible on the desktop regardless of Explorer settings.
        /// </summary>
        public static void Stage(string originalPath)
        {
            try
            {
                Directory.CreateDirectory(StagingDir);
                var staged = GetStagedPath(originalPath);
                bool stagedExists = File.Exists(staged) || Directory.Exists(staged);
                bool originalExists = File.Exists(originalPath) || Directory.Exists(originalPath);

                if (stagedExists && originalExists)
                {
                    // Duplicate — file is in staging AND on desktop (e.g. OneDrive
                    // restored it, or installer recreated a shortcut). Remove the
                    // desktop copy since the staged version is the one we manage.
                    try
                    {
                        if (File.Exists(originalPath))
                            File.Delete(originalPath);
                        else if (Directory.Exists(originalPath))
                            Directory.Delete(originalPath, true);
                    }
                    catch { }
                    return;
                }

                if (stagedExists)
                    return; // already staged, no desktop copy

                if (Directory.Exists(originalPath))
                    Directory.Move(originalPath, staged);
                else if (File.Exists(originalPath))
                    File.Move(originalPath, staged);
            }
            catch { }
        }

        /// <summary>
        /// Move a file/folder back from staging to its original desktop location.
        /// </summary>
        public static void Unstage(string originalPath)
        {
            try
            {
                var staged = GetStagedPath(originalPath);
                if (Directory.Exists(staged))
                    Directory.Move(staged, originalPath);
                else if (File.Exists(staged))
                    File.Move(staged, originalPath);
            }
            catch { }
        }
    }
}
