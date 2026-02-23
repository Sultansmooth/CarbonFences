using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace CarbonZones.Win32
{
    public class DesktopClickHook : IDisposable
    {
        private const int WH_MOUSE_LL = 14;
        private const int WM_LBUTTONDOWN = 0x0201;
        private const int WM_LBUTTONUP = 0x0202;
        private const int WM_MOUSEMOVE = 0x0200;

        private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int x;
            public int y;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MSLLHOOKSTRUCT
        {
            public POINT pt;
            public uint mouseData;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll")]
        private static extern IntPtr WindowFromPoint(POINT pt);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

        private const int VK_MENU = 0x12; // Alt key

        private IntPtr hookId = IntPtr.Zero;
        private readonly LowLevelMouseProc proc;

        // Double-click tracking
        private DateTime lastClickTime = DateTime.MinValue;
        private POINT lastClickPos;

        // Drag tracking (only when Alt is held)
        private POINT mouseDownPos;
        private bool mouseIsDown;
        private bool isDragging;
        private bool altWasHeld;

        public event EventHandler DesktopDoubleClicked;
        public event EventHandler<Rectangle> DesktopDragCompleted;

        public DesktopClickHook()
        {
            proc = HookCallback;
            using var curProcess = Process.GetCurrentProcess();
            using var curModule = curProcess.MainModule;
            hookId = SetWindowsHookEx(WH_MOUSE_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
        }

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                var hookStruct = Marshal.PtrToStructure<MSLLHOOKSTRUCT>(lParam);

                if (wParam == (IntPtr)WM_LBUTTONDOWN)
                {
                    if (IsDesktopWindow(hookStruct.pt))
                    {
                        var now = DateTime.Now;
                        var elapsed = (now - lastClickTime).TotalMilliseconds;
                        var dblTime = SystemInformation.DoubleClickTime;
                        var dblSize = SystemInformation.DoubleClickSize;

                        if (elapsed <= dblTime
                            && Math.Abs(hookStruct.pt.x - lastClickPos.x) <= dblSize.Width / 2
                            && Math.Abs(hookStruct.pt.y - lastClickPos.y) <= dblSize.Height / 2)
                        {
                            DesktopDoubleClicked?.Invoke(this, EventArgs.Empty);
                            lastClickTime = DateTime.MinValue;
                            mouseIsDown = false;
                            isDragging = false;
                        }
                        else
                        {
                            lastClickTime = now;
                            lastClickPos = hookStruct.pt;

                            // Only track for fence drawing when Alt is held
                            altWasHeld = (GetAsyncKeyState(VK_MENU) & 0x8000) != 0;
                            mouseDownPos = hookStruct.pt;
                            mouseIsDown = altWasHeld;
                            isDragging = false;
                        }
                    }
                    else
                    {
                        mouseIsDown = false;
                        isDragging = false;
                    }
                }
                else if (wParam == (IntPtr)WM_MOUSEMOVE && mouseIsDown && !isDragging)
                {
                    if (Math.Abs(hookStruct.pt.x - mouseDownPos.x) > SystemInformation.DragSize.Width ||
                        Math.Abs(hookStruct.pt.y - mouseDownPos.y) > SystemInformation.DragSize.Height)
                    {
                        if (IsDesktopWindow(hookStruct.pt))
                            isDragging = true;
                        else
                            mouseIsDown = false;
                    }
                }
                else if (wParam == (IntPtr)WM_LBUTTONUP)
                {
                    if (isDragging)
                    {
                        var rect = MakeRect(mouseDownPos, hookStruct.pt);
                        if (rect.Width > 30 && rect.Height > 30)
                            DesktopDragCompleted?.Invoke(this, rect);
                    }
                    mouseIsDown = false;
                    isDragging = false;
                }
            }
            return CallNextHookEx(hookId, nCode, wParam, lParam);
        }

        private static Rectangle MakeRect(POINT a, POINT b)
        {
            int x = Math.Min(a.x, b.x);
            int y = Math.Min(a.y, b.y);
            int w = Math.Abs(a.x - b.x);
            int h = Math.Abs(a.y - b.y);
            return new Rectangle(x, y, w, h);
        }

        private static bool IsDesktopWindow(POINT pt)
        {
            var hwnd = WindowFromPoint(pt);
            var sb = new StringBuilder(64);
            GetClassName(hwnd, sb, sb.Capacity);
            var className = sb.ToString();
            return className == "Progman" || className == "WorkerW"
                || className == "SysListView32" || className == "SHELLDLL_DefView";
        }

        public void Dispose()
        {
            if (hookId != IntPtr.Zero)
            {
                UnhookWindowsHookEx(hookId);
                hookId = IntPtr.Zero;
            }
        }
    }
}
