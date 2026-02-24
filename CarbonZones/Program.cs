using CarbonZones.Model;
using System;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using CarbonZones.Win32;

namespace CarbonZones
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            WindowUtil.SetPreferredAppMode(1);

            using (var mutex = new Mutex(true, "Carbon_Zones", out var createdNew))
            {
                if (createdNew)
                {
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);

                    FenceManager.Instance.LoadFences();
                    FenceManager.Instance.HideAllFencedIcons();
                    if (Application.OpenForms.Count == 0)
                        FenceManager.Instance.CreateFence("First fence");

                    using var hook = new DesktopClickHook();
                    hook.DesktopDoubleClicked += (s, e) => FenceManager.Instance.ToggleFences();

                    hook.DesktopDragCompleted += (s, rect) =>
                    {
                        var dialog = new EditDialog("New fence");
                        dialog.TopMost = true;
                        if (dialog.ShowDialog() == DialogResult.OK)
                            FenceManager.Instance.CreateFence(dialog.NewName, rect);
                    };

                    // System tray icon
                    var trayMenu = new ContextMenuStrip();
                    trayMenu.Items.Add("Show/Hide Fences", null, (s, e) => FenceManager.Instance.ToggleFences());
                    trayMenu.Items.Add("New Fence", null, (s, e) => FenceManager.Instance.CreateFence("New fence"));
                    trayMenu.Items.Add(new ToolStripSeparator());
                    trayMenu.Items.Add("Exit", null, (s, e) => Application.Exit());

                    var appIcon = LoadAppIcon();
                    using var trayIcon = new NotifyIcon
                    {
                        Icon = appIcon ?? SystemIcons.Application,
                        Text = "Carbon Zones",
                        Visible = true,
                        ContextMenuStrip = trayMenu
                    };
                    trayIcon.DoubleClick += (s, e) => FenceManager.Instance.ToggleFences();

                    Application.ApplicationExit += (s, e) => FenceManager.Instance.UnhideAllDesktopIcons();

                    Application.Run();
                }
            }
        }
        private static Icon LoadAppIcon()
        {
            var exeDir = Path.GetDirectoryName(Application.ExecutablePath);
            var iconPath = Path.Combine(exeDir, "icon.ico");
            if (File.Exists(iconPath))
                return new Icon(iconPath);
            return null;
        }
    }
}
