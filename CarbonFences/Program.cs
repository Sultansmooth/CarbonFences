using CarbonFences.Model;
using System;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using CarbonFences.Win32;

namespace CarbonFences
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            WindowUtil.SetPreferredAppMode(1);

            using (var mutex = new Mutex(true, "Carbon_Fences", out var createdNew))
            {
                if (createdNew)
                {
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);

                    FenceManager.Instance.LoadFences();
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

                    using var trayIcon = new NotifyIcon
                    {
                        Icon = SystemIcons.Application,
                        Text = "Carbon Fences",
                        Visible = true,
                        ContextMenuStrip = trayMenu
                    };
                    trayIcon.DoubleClick += (s, e) => FenceManager.Instance.ToggleFences();

                    Application.Run();
                }
            }
        }
    }
}
