using CarbonFences.Model;
using CarbonFences.Util;
using CarbonFences.Win32;
using Peter;
using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using static CarbonFences.Win32.WindowUtil;

namespace CarbonFences
{
    public partial class FenceWindow : Form
    {
        protected override bool ShowWithoutActivation => true;

        private int logicalTitleHeight;
        private int titleHeight;
        private const int titleOffset = 3;
        private const int itemWidth = 75;
        private const int itemHeight = 32 + itemPadding + textHeight;
        private const int textHeight = 35;
        private const int itemPadding = 15;
        private const float shadowDist = 1.5f;

        private readonly FenceInfo fenceInfo;

        private Font titleFont;
        private Font iconFont;

        private string selectedItem;
        private string hoveringItem;
        private bool shouldUpdateSelection;
        private bool shouldRunDoubleClick;
        private bool hasSelectionUpdated;
        private bool hasHoverUpdated;
        private bool isMinified;
        private bool isMouseOver;
        private float hoverAlpha;
        private int prevHeight;

        private int scrollHeight;
        private int scrollOffset;

        private Point dragStartPoint;
        private string dragOutItem;

        private readonly ThrottledExecution throttledMove = new ThrottledExecution(TimeSpan.FromSeconds(4));
        private readonly ThrottledExecution throttledResize = new ThrottledExecution(TimeSpan.FromSeconds(4));

        private readonly ShellContextMenu shellContextMenu = new ShellContextMenu();

        private readonly ThumbnailProvider thumbnailProvider = new ThumbnailProvider();

        private readonly System.Windows.Forms.Timer hoverTimer = new System.Windows.Forms.Timer { Interval = 16 };

        private void ReloadFonts()
        {
            var family = new FontFamily("Segoe UI");
            titleFont = new Font(family, (int)Math.Floor(logicalTitleHeight / 2.0));
            iconFont = new Font(family, 9);
        }

        public FenceWindow(FenceInfo fenceInfo)
        {
            InitializeComponent();
            DropShadow.ApplyShadows(this);
            BlurUtil.EnableBlur(Handle);
            WindowUtil.HideFromAltTab(Handle);
            logicalTitleHeight = (fenceInfo.TitleHeight < 16 || fenceInfo.TitleHeight > 100) ? 35 : fenceInfo.TitleHeight;
            titleHeight = LogicalToDeviceUnits(logicalTitleHeight);

            this.MouseWheel += FenceWindow_MouseWheel;
            this.MouseDown += FenceWindow_MouseDown;
            thumbnailProvider.IconThumbnailLoaded += ThumbnailProvider_IconThumbnailLoaded;
            hoverTimer.Tick += HoverTimer_Tick;

            ReloadFonts();

            AllowDrop = true;

            this.fenceInfo = fenceInfo;
            Text = fenceInfo.Name;
            Location = new Point(fenceInfo.PosX, fenceInfo.PosY);

            Width = fenceInfo.Width;
            Height = fenceInfo.Height;

            prevHeight = Height;
            lockedToolStripMenuItem.Checked = fenceInfo.Locked;
            minifyToolStripMenuItem.Checked = fenceInfo.CanMinify;
            Minify();

            FenceManager.Instance.RegisterWindow(this);
        }

        private const int WM_RBUTTONUP = 0x0205;
        private const int WM_NCRBUTTONUP = 0x00A5;
        private const int WM_MOUSEACTIVATE = 0x0021;
        private const int MA_NOACTIVATE = 3;
        protected override void WndProc(ref Message m)
        {
            // Remove border
            if (m.Msg == 0x0083)
            {
                m.Result = IntPtr.Zero;
                return;
            }

            // Mouse leave
            var myrect = new Rectangle(Location, Size);
            if (m.Msg == 0x02a2 && !myrect.IntersectsWith(new Rectangle(MousePosition, new Size(1, 1))))
            {
                Minify();
            }

            // Prevent maximize
            if ((m.Msg == WM_SYSCOMMAND) && m.WParam.ToInt32() == 0xF032)
            {
                m.Result = IntPtr.Zero;
                return;
            }

            // Prevent activation — fence should never steal focus
            if (m.Msg == WM_MOUSEACTIVATE)
            {
                m.Result = (IntPtr)MA_NOACTIVATE;
                return;
            }

            if (m.Msg == WM_SETFOCUS)
            {
                m.Result = IntPtr.Zero;
                return;
            }

            // Right-click context menu — handle directly for reliability
            if (m.Msg == WM_RBUTTONUP)
            {
                var pt = new Point(
                    (short)(m.LParam.ToInt64() & 0xFFFF),
                    (short)(m.LParam.ToInt64() >> 16 & 0xFFFF));
                Invalidate(); // update hover state
                if (hoveringItem != null && !ModifierKeys.HasFlag(Keys.Shift))
                    shellContextMenu.ShowContextMenu(new[] { new FileInfo(hoveringItem) }, PointToScreen(pt));
                else
                    appContextMenu.Show(this, pt);
                return;
            }

            // Right-click on title bar (non-client caption area)
            if (m.Msg == WM_NCRBUTTONUP)
            {
                var screenPt = new Point(
                    (short)(m.LParam.ToInt64() & 0xFFFF),
                    (short)(m.LParam.ToInt64() >> 16 & 0xFFFF));
                appContextMenu.Show(screenPt);
                return;
            }

            // Other messages
            base.WndProc(ref m);

            // If not locked and using the left mouse button
            if (MouseButtons == MouseButtons.Right || lockedToolStripMenuItem.Checked)
                return;

            // Then, allow dragging and resizing
            if (m.Msg == WM_NCHITTEST)
            {
                var pt = PointToClient(new Point(m.LParam.ToInt32()));

                if ((int)m.Result == HTCLIENT && pt.Y < titleHeight)
                {
                    m.Result = (IntPtr)HTCAPTION;
                    FenceWindow_MouseEnter(null, null);
                }

                if (pt.X < 10 && pt.Y < 10)
                    m.Result = new IntPtr(HTTOPLEFT);
                else if (pt.X > (Width - 10) && pt.Y < 10)
                    m.Result = new IntPtr(HTTOPRIGHT);
                else if (pt.X < 10 && pt.Y > (Height - 10))
                    m.Result = new IntPtr(HTBOTTOMLEFT);
                else if (pt.X > (Width - 10) && pt.Y > (Height - 10))
                    m.Result = new IntPtr(HTBOTTOMRIGHT);
                else if (pt.Y > (Height - 10))
                    m.Result = new IntPtr(HTBOTTOM);
                else if (pt.X < 10)
                    m.Result = new IntPtr(HTLEFT);
                else if (pt.X > (Width - 10))
                    m.Result = new IntPtr(HTRIGHT);
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(this, "Really remove this fence?", "Remove", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                FenceManager.Instance.RemoveFence(fenceInfo);
                Close();
            }
        }

        private void deleteItemToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowDesktopIcon(hoveringItem);
            fenceInfo.Files.Remove(hoveringItem);
            hoveringItem = null;
            Save();
            Refresh();
        }

        private void contextMenuStrip1_Opening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            deleteItemToolStripMenuItem.Visible = hoveringItem != null;
        }

        private void FenceWindow_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop) && !lockedToolStripMenuItem.Checked)
                e.Effect = DragDropEffects.Move;
        }

        private void FenceWindow_DragDrop(object sender, DragEventArgs e)
        {
            var dropped = (string[])e.Data.GetData(DataFormats.FileDrop);
            foreach (var file in dropped)
            {
                if (!fenceInfo.Files.Contains(file) && ItemExists(file))
                {
                    fenceInfo.Files.Add(file);
                    HideDesktopIcon(file);
                }
            }
            Save();
            Refresh();
        }

        private void FenceWindow_Resize(object sender, EventArgs e)
        {
            throttledResize.Run(() =>
            {
                fenceInfo.Width = Width;
                fenceInfo.Height = isMinified ? prevHeight : Height;
                Save();
            });

            Invalidate();
        }

        private void FenceWindow_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                dragStartPoint = e.Location;
                dragOutItem = hoveringItem;
            }
        }

        private void FenceWindow_MouseMove(object sender, MouseEventArgs e)
        {
            // Drag item out of fence
            if (e.Button == MouseButtons.Left && dragOutItem != null && !lockedToolStripMenuItem.Checked)
            {
                if (Math.Abs(e.X - dragStartPoint.X) > SystemInformation.DragSize.Width ||
                    Math.Abs(e.Y - dragStartPoint.Y) > SystemInformation.DragSize.Height)
                {
                    var data = new DataObject(DataFormats.FileDrop, new[] { dragOutItem });
                    var itemPath = dragOutItem;
                    dragOutItem = null;
                    DoDragDrop(data, DragDropEffects.Move | DragDropEffects.Copy);

                    // If mouse ended outside the fence, remove item
                    var mousePos = PointToClient(Control.MousePosition);
                    if (!ClientRectangle.Contains(mousePos))
                    {
                        ShowDesktopIcon(itemPath);
                        fenceInfo.Files.Remove(itemPath);
                        Save();
                    }
                    Refresh();
                    return;
                }
            }
            Invalidate();
        }

        private void HoverTimer_Tick(object sender, EventArgs e)
        {
            float target = isMouseOver ? 1.0f : 0.0f;
            float step = 0.08f;

            if (Math.Abs(hoverAlpha - target) < step)
            {
                hoverAlpha = target;
                hoverTimer.Stop();
            }
            else
            {
                hoverAlpha += (target > hoverAlpha) ? step : -step;
            }

            Invalidate();
        }

        private void FenceWindow_MouseEnter(object sender, EventArgs e)
        {
            isMouseOver = true;
            hoverTimer.Start();

            if (minifyToolStripMenuItem.Checked && isMinified)
            {
                isMinified = false;
                Height = prevHeight;
            }
        }

        private void FenceWindow_MouseLeave(object sender, EventArgs e)
        {
            isMouseOver = false;
            hoverTimer.Start();

            Minify();
            selectedItem = null;
            Refresh();
        }

        private void Minify()
        {
            if (minifyToolStripMenuItem.Checked && !isMinified)
            {
                isMinified = true;
                prevHeight = Height;
                Height = titleHeight;
                Refresh();
            }
        }

        private void minifyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (isMinified)
            {
                Height = prevHeight;
                isMinified = false;
            }
            fenceInfo.CanMinify = minifyToolStripMenuItem.Checked;
            Save();

        }

        private void FenceWindow_Click(object sender, EventArgs e)
        {
            shouldUpdateSelection = true;
            Refresh();
        }

        private void FenceWindow_DoubleClick(object sender, EventArgs e)
        {
            shouldRunDoubleClick = true;
            Refresh();
        }

        private void FenceWindow_Paint(object sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            g.Clip = new Region(ClientRectangle);
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

            // Background — very transparent when idle, solid on hover
            int bgAlpha = (int)(30 + 120 * hoverAlpha);
            using (var bgBrush = new SolidBrush(Color.FromArgb(bgAlpha, Color.Black)))
                g.FillRectangle(bgBrush, ClientRectangle);

            // Hover border glow
            if (hoverAlpha > 0)
            {
                int glowAlpha = (int)(60 * hoverAlpha);
                using var borderPen = new Pen(Color.FromArgb(glowAlpha, 120, 180, 255), 1);
                g.DrawRectangle(borderPen, 0, 0, Width - 1, Height - 1);
            }

            // Title background, then text on top
            int titleAlpha = (int)(15 + 60 * hoverAlpha);
            using (var titleBgBrush = new SolidBrush(Color.FromArgb(titleAlpha, Color.Black)))
                g.FillRectangle(titleBgBrush, new RectangleF(0, 0, Width, titleHeight));
            int textAlpha = (int)(60 + 195 * hoverAlpha);
            using (var titleBrush = new SolidBrush(Color.FromArgb(textAlpha, 255, 255, 255)))
            using (var titleFormat = new StringFormat { Alignment = StringAlignment.Center })
                g.DrawString(Text, titleFont, titleBrush, new PointF(Width / 2, titleOffset), titleFormat);

            // Items
            var x = itemPadding;
            var y = itemPadding;
            scrollHeight = 0;
            g.Clip = new Region(new Rectangle(0, titleHeight, Width, Height - titleHeight));
            foreach (var file in fenceInfo.Files)
            {
                var entry = FenceEntry.FromPath(file);
                if (entry == null)
                    continue;

                RenderEntry(g, entry, x, y + titleHeight - scrollOffset);

                var itemBottom = y + itemHeight;
                if (itemBottom > scrollHeight)
                    scrollHeight = itemBottom;

                x += itemWidth + itemPadding;
                if (x + itemWidth > Width)
                {
                    x = itemPadding;
                    y += itemHeight + itemPadding;
                }
            }

            scrollHeight -= (ClientRectangle.Height - titleHeight);

            // Scroll bars
            if (scrollHeight > 0)
            {
                var contentHeight = Height - titleHeight;
                var scrollbarHeight = contentHeight - scrollHeight;
                using (var scrollBrush = new SolidBrush(Color.FromArgb(150, Color.Black)))
                    g.FillRectangle(scrollBrush, new Rectangle(Width - 5, titleHeight + scrollOffset, 5, scrollbarHeight));

                scrollOffset = Math.Min(scrollOffset, scrollHeight);
            }

            // Click handlers
            if (shouldUpdateSelection && !hasSelectionUpdated)
                selectedItem = null;

            if (!hasHoverUpdated)
                hoveringItem = null;

            shouldRunDoubleClick = false;
            shouldUpdateSelection = false;
            hasSelectionUpdated = false;
            hasHoverUpdated = false;
        }

        private void RenderEntry(Graphics g, FenceEntry entry, int x, int y)
        {
            var icon = entry.ExtractIcon(thumbnailProvider);
            var name = entry.Name;

            var textPosition = new PointF(x, y + icon.Height + 5);
            var textMaxSize = new SizeF(itemWidth, textHeight);

            using var stringFormat = new StringFormat { Alignment = StringAlignment.Center, Trimming = StringTrimming.EllipsisCharacter };

            var textSize = g.MeasureString(name, iconFont, textMaxSize, stringFormat);
            var outlineRect = new Rectangle(x - 2, y - 2, itemWidth + 2, icon.Height + (int)textSize.Height + 5 + 2);
            var outlineRectInner = outlineRect.Shrink(1);

            var mousePos = PointToClient(MousePosition);
            var mouseOver = mousePos.X >= x && mousePos.Y >= y && mousePos.X < x + outlineRect.Width && mousePos.Y < y + outlineRect.Height;

            if (mouseOver)
            {
                hoveringItem = entry.Path;
                hasHoverUpdated = true;
            }

            if (mouseOver && shouldUpdateSelection)
            {
                selectedItem = entry.Path;
                shouldUpdateSelection = false;
                hasSelectionUpdated = true;
            }

            if (mouseOver && shouldRunDoubleClick)
            {
                shouldRunDoubleClick = false;
                entry.Open();
            }

            if (selectedItem == entry.Path)
            {
                using var borderPen = new Pen(Color.FromArgb(120, SystemColors.ActiveBorder));
                g.DrawRectangle(borderPen, outlineRectInner);
                if (mouseOver)
                {
                    using var fillBrush = new SolidBrush(Color.FromArgb(100, SystemColors.GradientActiveCaption));
                    g.FillRectangle(fillBrush, outlineRect);
                }
                else
                {
                    using var fillBrush = new SolidBrush(Color.FromArgb(80, SystemColors.GradientInactiveCaption));
                    g.FillRectangle(fillBrush, outlineRect);
                }
            }
            else if (mouseOver)
            {
                using var borderPen = new Pen(Color.FromArgb(120, SystemColors.ActiveBorder));
                using var fillBrush = new SolidBrush(Color.FromArgb(80, SystemColors.ActiveCaption));
                g.DrawRectangle(borderPen, outlineRectInner);
                g.FillRectangle(fillBrush, outlineRect);
            }

            g.DrawIcon(icon, x + itemWidth / 2 - icon.Width / 2, y);
            using (var shadowBrush = new SolidBrush(Color.FromArgb(180, 15, 15, 15)))
                g.DrawString(name, iconFont, shadowBrush, new RectangleF(textPosition.Move(shadowDist, shadowDist), textMaxSize), stringFormat);
            g.DrawString(name, iconFont, Brushes.White, new RectangleF(textPosition, textMaxSize), stringFormat);
        }

        private void renameToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var dialog = new EditDialog(Text);
            if (dialog.ShowDialog(this) == DialogResult.OK)
            {
                Text = dialog.NewName;
                fenceInfo.Name = Text;
                Refresh();
                Save();
            }
        }

        private void newFenceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FenceManager.Instance.CreateFence("New fence");
        }

        private void FenceWindow_FormClosed(object sender, FormClosedEventArgs e)
        {
            FenceManager.Instance.UnregisterWindow(this);
            // App stays alive via tray icon even with no fences open
        }

        private readonly object saveLock = new object();
        private void Save()
        {
            lock (saveLock)
            {
                FenceManager.Instance.UpdateFence(fenceInfo);
            }
        }

        private void FenceWindow_LocationChanged(object sender, EventArgs e)
        {
            throttledMove.Run(() =>
            {
                fenceInfo.PosX = Location.X;
                fenceInfo.PosY = Location.Y;
                Save();
            });
        }

        private void lockedToolStripMenuItem_Click(object sender, EventArgs e)
        {
            fenceInfo.Locked = lockedToolStripMenuItem.Checked;
            Save();
        }

        private void FenceWindow_Load(object sender, EventArgs e)
        {

        }

        private void titleSizeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var dialog = new HeightDialog(fenceInfo.TitleHeight);
            if (dialog.ShowDialog(this) == DialogResult.OK)
            {
                fenceInfo.TitleHeight = dialog.TitleHeight;
                logicalTitleHeight = dialog.TitleHeight;
                titleHeight = LogicalToDeviceUnits(logicalTitleHeight);
                ReloadFonts();
                Minify();
                if (isMinified)
                {
                    Height = titleHeight;
                }
                Refresh();
                Save();
            }
        }

        private void FenceWindow_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right)
                return;

            if (hoveringItem != null && !ModifierKeys.HasFlag(Keys.Shift))
            {
                shellContextMenu.ShowContextMenu(new[] { new FileInfo(hoveringItem) }, MousePosition);
            }
            else
            {
                appContextMenu.Show(this, e.Location);
            }
        }

        private void FenceWindow_MouseWheel(object sender, MouseEventArgs e)
        {
            if (scrollHeight < 1)
                return;

            scrollOffset -= Math.Sign(e.Delta) * 10;
            if (scrollOffset < 0)
                scrollOffset = 0;
            if (scrollOffset > scrollHeight)
                scrollOffset = scrollHeight;

            Invalidate();
        }

        private void ThumbnailProvider_IconThumbnailLoaded(object sender, EventArgs e)
        {
            Invalidate();
        }

        private bool ItemExists(string path)
        {
            return File.Exists(path) || Directory.Exists(path);
        }

        [DllImport("shell32.dll")]
        private static extern void SHChangeNotify(int wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);
        private const int SHCNE_UPDATEITEM = 0x00002000;
        private const uint SHCNF_PATHW = 0x0005;

        private static void NotifyShell(string path)
        {
            var ptr = Marshal.StringToHGlobalUni(path);
            try { SHChangeNotify(SHCNE_UPDATEITEM, SHCNF_PATHW, ptr, IntPtr.Zero); }
            finally { Marshal.FreeHGlobal(ptr); }
        }

        private static void HideDesktopIcon(string path)
        {
            if (!IsDesktopPath(path)) return;
            try
            {
                var attrs = File.GetAttributes(path);
                if (!attrs.HasFlag(FileAttributes.Hidden))
                    File.SetAttributes(path, attrs | FileAttributes.Hidden);
                NotifyShell(path);
            }
            catch { }
        }

        private static void ShowDesktopIcon(string path)
        {
            if (!IsDesktopPath(path)) return;
            try
            {
                var attrs = File.GetAttributes(path);
                if (attrs.HasFlag(FileAttributes.Hidden))
                    File.SetAttributes(path, attrs & ~FileAttributes.Hidden);
                NotifyShell(path);
            }
            catch { }
        }

        private static bool IsDesktopPath(string path)
        {
            var dir = Path.GetDirectoryName(path);
            if (dir == null) return false;
            var userDesktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            var publicDesktop = Environment.GetFolderPath(Environment.SpecialFolder.CommonDesktopDirectory);
            return string.Equals(dir, userDesktop, StringComparison.OrdinalIgnoreCase)
                || string.Equals(dir, publicDesktop, StringComparison.OrdinalIgnoreCase);
        }
    }

}

