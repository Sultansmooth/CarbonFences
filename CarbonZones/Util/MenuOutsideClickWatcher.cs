using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace CarbonZones.Util
{
    /// <summary>
    /// Wires a ContextMenuStrip so that:
    ///  1. On Opened: installs a low-level system mouse hook so clicks
    ///     anywhere outside the menu's screen-rect (including the desktop
    ///     or other apps) close it. Standard auto-close is unreliable when
    ///     the menu's owner form is non-activating (MA_NOACTIVATE).
    ///  2. On Opened: applies a rounded Region to the menu window so the
    ///     custom renderer's rounded corners show against any background.
    ///     Setting the Region inside the paint loop caused paint storms,
    ///     which made hover state intermittently fail to render.
    ///  3. On Closed: removes the hook to keep the system-wide hook overhead
    ///     scoped to the brief lifetime of the open menu.
    /// </summary>
    public static class MenuOutsideClickWatcher
    {
        public static void Attach(ToolStripDropDown menu, int cornerRadius = 8)
        {
            IntPtr hookHandle = IntPtr.Zero;
            // Captured by both lambdas below — must remain rooted while the
            // hook is installed, otherwise the GC will collect the delegate
            // and the OS will call into freed memory.
            HookProc proc = null;
            proc = (code, wParam, lParam) =>
            {
                if (code >= 0)
                {
                    int msg = wParam.ToInt32();
                    if (msg == WM_LBUTTONDOWN || msg == WM_RBUTTONDOWN || msg == WM_MBUTTONDOWN ||
                        msg == WM_NCLBUTTONDOWN || msg == WM_NCRBUTTONDOWN)
                    {
                        var data = Marshal.PtrToStructure<MSLLHOOKSTRUCT>(lParam);
                        var pt = new Point(data.pt.x, data.pt.y);
                        if (menu.Visible && !menu.Bounds.Contains(pt))
                        {
                            try { menu.BeginInvoke(new Action(() => { if (menu.Visible) menu.Close(); })); }
                            catch { }
                        }
                    }
                }
                return CallNextHookEx(hookHandle, code, wParam, lParam);
            };

            menu.Opened += (s, e) =>
            {
                ApplyRoundedRegion(menu, cornerRadius);
                try
                {
                    var hMod = GetModuleHandle(null); // any non-null module handle works for WH_MOUSE_LL
                    hookHandle = SetWindowsHookEx(WH_MOUSE_LL, proc, hMod, 0);
                }
                catch { hookHandle = IntPtr.Zero; }
            };

            menu.Closed += (s, e) =>
            {
                if (hookHandle != IntPtr.Zero)
                {
                    try { UnhookWindowsHookEx(hookHandle); } catch { }
                    hookHandle = IntPtr.Zero;
                }
                // Drop the cached region so the next Open re-measures (size
                // can change if items are added/removed between opens).
                try { menu.Region = null; } catch { }
            };
        }

        private static void ApplyRoundedRegion(ToolStripDropDown menu, int radius)
        {
            try
            {
                int w = menu.Width;
                int h = menu.Height;
                if (w <= 0 || h <= 0) return;

                int d = radius * 2;
                if (d > w) d = w;
                if (d > h) d = h;

                using var path = new GraphicsPath();
                if (d <= 0)
                {
                    path.AddRectangle(new Rectangle(0, 0, w, h));
                }
                else
                {
                    path.AddArc(0, 0, d, d, 180, 90);
                    path.AddArc(w - d, 0, d, d, 270, 90);
                    path.AddArc(w - d, h - d, d, d, 0, 90);
                    path.AddArc(0, h - d, d, d, 90, 90);
                    path.CloseFigure();
                }
                menu.Region = new Region(path);
            }
            catch { }
        }

        // --- P/Invoke ---
        private const int WH_MOUSE_LL = 14;
        private const int WM_LBUTTONDOWN = 0x0201;
        private const int WM_RBUTTONDOWN = 0x0204;
        private const int WM_MBUTTONDOWN = 0x0207;
        private const int WM_NCLBUTTONDOWN = 0x00A1;
        private const int WM_NCRBUTTONDOWN = 0x00A4;

        private delegate IntPtr HookProc(int code, IntPtr wParam, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        private struct MSLLHOOKSTRUCT { public POINT pt; public uint mouseData; public uint flags; public uint time; public IntPtr dwExtraInfo; }

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT { public int x; public int y; }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);
    }
}
