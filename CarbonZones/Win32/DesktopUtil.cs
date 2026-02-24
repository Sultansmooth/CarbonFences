using System;
using System.Runtime.InteropServices;
using System.Text;

namespace CarbonZones.Win32
{
    public class DesktopUtil
    {
        private const Int32 GWL_STYLE = -16;
        private const Int32 GWL_HWNDPARENT = -8;
        private const Int32 WS_MAXIMIZEBOX = 0x00010000;
        private const Int32 WS_MINIMIZEBOX = 0x00020000;
        private const int SW_HIDE = 0;
        private const int SW_SHOW = 5;

        [DllImport("User32.dll", EntryPoint = "GetWindowLong")]
        private extern static Int32 GetWindowLongPtr(IntPtr hWnd, Int32 nIndex);

        [DllImport("User32.dll", EntryPoint = "SetWindowLong")]
        private extern static Int32 SetWindowLongPtr(IntPtr hWnd, Int32 nIndex, Int32 dwNewLong);

        [DllImport("user32.dll", SetLastError = true)]
        static extern int SetWindowLong(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpWindowClass, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string className, string windowTitle);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool IsWindowVisible(IntPtr hWnd);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        public static void PreventMinimize(IntPtr handle)
        {
            Int32 windowStyle = GetWindowLongPtr(handle, GWL_STYLE);
            SetWindowLongPtr(handle, GWL_STYLE, windowStyle & ~WS_MAXIMIZEBOX & ~WS_MINIMIZEBOX);
        }

        public static void GlueToDesktop(IntPtr handle)
        {
            IntPtr nWinHandle = FindWindowEx(IntPtr.Zero, IntPtr.Zero, "Progman", null);
            SetWindowLongPtr(handle, GWL_HWNDPARENT, nWinHandle.ToInt32());
        }

        [DllImport("user32.dll")]
        static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        private const uint LVM_ARRANGE = 0x1016;
        private const int LVA_DEFAULT = 0x0000;

        public static void SetDesktopIconsVisible(bool visible)
        {
            IntPtr defView = GetShellDefView();
            if (defView == IntPtr.Zero) return;

            IntPtr listView = FindWindowEx(defView, IntPtr.Zero, "SysListView32", null);
            if (listView == IntPtr.Zero) return;

            ShowWindow(listView, visible ? SW_SHOW : SW_HIDE);
        }

        /// <summary>
        /// Sends LVM_ARRANGE to the desktop ListView to force it to re-read item states.
        /// </summary>
        public static void RefreshDesktopIcons()
        {
            IntPtr defView = GetShellDefView();
            if (defView == IntPtr.Zero) return;

            IntPtr listView = FindWindowEx(defView, IntPtr.Zero, "SysListView32", null);
            if (listView == IntPtr.Zero) return;

            SendMessage(listView, LVM_ARRANGE, (IntPtr)LVA_DEFAULT, IntPtr.Zero);
        }

        private static IntPtr GetShellDefView()
        {
            IntPtr progman = FindWindow("Progman", null);
            IntPtr defView = FindWindowEx(progman, IntPtr.Zero, "SHELLDLL_DefView", null);
            if (defView != IntPtr.Zero) return defView;

            IntPtr result = IntPtr.Zero;
            EnumWindows((hwnd, lParam) =>
            {
                var sb = new StringBuilder(256);
                GetClassName(hwnd, sb, 256);
                if (sb.ToString() == "WorkerW")
                {
                    IntPtr shell = FindWindowEx(hwnd, IntPtr.Zero, "SHELLDLL_DefView", null);
                    if (shell != IntPtr.Zero)
                    {
                        result = shell;
                        return false;
                    }
                }
                return true;
            }, IntPtr.Zero);
            return result;
        }
    }
}