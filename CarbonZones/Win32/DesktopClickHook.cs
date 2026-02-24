using System;
using System.Collections.Generic;
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

        // Fence window handle registry — prevents desktop-click detection inside fences
        private static readonly HashSet<IntPtr> fenceHandles = new HashSet<IntPtr>();
        public static void RegisterFenceHandle(IntPtr hwnd) => fenceHandles.Add(hwnd);
        public static void UnregisterFenceHandle(IntPtr hwnd) => fenceHandles.Remove(hwnd);

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

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left, Top, Right, Bottom;
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string className, string windowTitle);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("kernel32.dll")]
        private static extern IntPtr OpenProcess(uint dwDesiredAccess, bool bInheritHandle, uint dwProcessId);

        [DllImport("kernel32.dll")]
        private static extern IntPtr VirtualAllocEx(IntPtr hProcess, IntPtr lpAddress, uint dwSize, uint flAllocationType, uint flProtect);

        [DllImport("kernel32.dll")]
        private static extern bool VirtualFreeEx(IntPtr hProcess, IntPtr lpAddress, uint dwSize, uint dwFreeType);

        [DllImport("kernel32.dll")]
        private static extern bool WriteProcessMemory(IntPtr hProcess, IntPtr lpBaseAddress, byte[] lpBuffer, uint nSize, out int lpNumberOfBytesWritten);

        [DllImport("kernel32.dll")]
        private static extern bool ReadProcessMemory(IntPtr hProcess, IntPtr lpBaseAddress, byte[] lpBuffer, uint nSize, out int lpNumberOfBytesRead);

        [DllImport("kernel32.dll")]
        private static extern bool CloseHandle(IntPtr hObject);

        private const uint LVM_HITTEST = 0x1012;
        private const uint PROCESS_VM_OPERATION = 0x0008;
        private const uint PROCESS_VM_READ = 0x0010;
        private const uint PROCESS_VM_WRITE = 0x0020;
        private const uint MEM_COMMIT = 0x1000;
        private const uint MEM_RELEASE = 0x8000;
        private const uint PAGE_READWRITE = 0x04;

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

            // If WindowFromPoint returned a registered fence handle, it's not the desktop
            if (fenceHandles.Contains(hwnd))
                return false;

            // Even if class name looks like desktop, check if point falls inside any fence window
            foreach (var fenceHwnd in fenceHandles)
            {
                if (IsWindowVisible(fenceHwnd) && GetWindowRect(fenceHwnd, out var rect))
                {
                    if (pt.x >= rect.Left && pt.x <= rect.Right &&
                        pt.y >= rect.Top && pt.y <= rect.Bottom)
                        return false;
                }
            }

            var sb = new StringBuilder(64);
            GetClassName(hwnd, sb, sb.Capacity);
            var className = sb.ToString();
            if (className != "Progman" && className != "WorkerW"
                && className != "SysListView32" && className != "SHELLDLL_DefView")
                return false;

            // It's the desktop — but only count as empty space if no icon was hit
            return !IsDesktopIconAtPoint(pt);
        }

        /// <summary>
        /// Uses LVM_HITTEST on the desktop SysListView32 to check if a screen
        /// point lands on an actual desktop icon. The ListView lives in
        /// explorer.exe, so we must allocate memory in that process.
        /// </summary>
        private static bool IsDesktopIconAtPoint(POINT screenPt)
        {
            try
            {
                IntPtr listView = GetDesktopListView();
                if (listView == IntPtr.Zero) return false;

                // Convert screen coords to ListView client coords
                GetWindowRect(listView, out var lvRect);
                int localX = screenPt.x - lvRect.Left;
                int localY = screenPt.y - lvRect.Top;

                // LVHITTESTINFO struct: POINT pt (8 bytes) + uint flags (4) + int iItem (4) + ...
                // We only need 16 bytes for the basic hit test
                const int structSize = 16;
                byte[] buffer = new byte[structSize];
                BitConverter.GetBytes(localX).CopyTo(buffer, 0);
                BitConverter.GetBytes(localY).CopyTo(buffer, 4);

                GetWindowThreadProcessId(listView, out uint pid);
                IntPtr hProcess = OpenProcess(PROCESS_VM_OPERATION | PROCESS_VM_READ | PROCESS_VM_WRITE, false, pid);
                if (hProcess == IntPtr.Zero) return false;

                try
                {
                    IntPtr remoteMem = VirtualAllocEx(hProcess, IntPtr.Zero, (uint)structSize, MEM_COMMIT, PAGE_READWRITE);
                    if (remoteMem == IntPtr.Zero) return false;

                    try
                    {
                        WriteProcessMemory(hProcess, remoteMem, buffer, (uint)structSize, out _);
                        SendMessage(listView, LVM_HITTEST, IntPtr.Zero, remoteMem);
                        ReadProcessMemory(hProcess, remoteMem, buffer, (uint)structSize, out _);

                        int iItem = BitConverter.ToInt32(buffer, 12);
                        return iItem >= 0; // >= 0 means an icon was hit
                    }
                    finally
                    {
                        VirtualFreeEx(hProcess, remoteMem, 0, MEM_RELEASE);
                    }
                }
                finally
                {
                    CloseHandle(hProcess);
                }
            }
            catch
            {
                return false; // On any error, assume empty space
            }
        }

        private static IntPtr GetDesktopListView()
        {
            IntPtr progman = FindWindow("Progman", null);
            IntPtr defView = FindWindowEx(progman, IntPtr.Zero, "SHELLDLL_DefView", null);
            if (defView == IntPtr.Zero)
            {
                EnumWindows((hwnd, lParam) =>
                {
                    var sb = new StringBuilder(64);
                    GetClassName(hwnd, sb, 64);
                    if (sb.ToString() == "WorkerW")
                    {
                        IntPtr shell = FindWindowEx(hwnd, IntPtr.Zero, "SHELLDLL_DefView", null);
                        if (shell != IntPtr.Zero)
                        {
                            defView = shell;
                            return false;
                        }
                    }
                    return true;
                }, IntPtr.Zero);
            }
            if (defView == IntPtr.Zero) return IntPtr.Zero;
            return FindWindowEx(defView, IntPtr.Zero, "SysListView32", null);
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
