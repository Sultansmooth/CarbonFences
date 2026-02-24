using CarbonZones.Model;
using CarbonZones.Util;
using CarbonZones.Win32;
using Peter;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using static CarbonZones.Win32.WindowUtil;

namespace CarbonZones
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

        // Tab state
        private int activeTabIndex = 0;
        private int hoveringTabIndex = -1;
        private int dragOverTabIndex = -1;
        private int draggingTabIndex = -1;
        private Point tabDragStartPoint;
        private bool tabWasDragged;
        private const int tabBarHeight = 24;
        private int scaledTabBarHeight;
        private const int tabPadding = 6;
        private const int tabSpacing = 1;
        private const int addButtonWidth = 24;
        private Font tabFont;
        private ContextMenuStrip tabContextMenu;
        private int contextMenuTabIndex = -1;

        private int headerHeight => titleHeight + scaledTabBarHeight;
        private List<string> ActiveFiles => fenceInfo.Tabs[activeTabIndex].Files;

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
            scaledTabBarHeight = LogicalToDeviceUnits(tabBarHeight);
            tabFont = new Font(new FontFamily("Segoe UI"), 8);

            // Tab context menu
            tabContextMenu = new ContextMenuStrip();
            tabContextMenu.Items.Add("Rename Tab", null, TabRename_Click);
            tabContextMenu.Items.Add("Delete Tab", null, TabDelete_Click);

            this.MouseWheel += FenceWindow_MouseWheel;
            this.MouseDown += FenceWindow_MouseDown;
            this.MouseUp += FenceWindow_MouseUp;
            thumbnailProvider.IconThumbnailLoaded += ThumbnailProvider_IconThumbnailLoaded;
            hoverTimer.Tick += HoverTimer_Tick;

            ReloadFonts();

            AllowDrop = true;

            this.fenceInfo = fenceInfo;

            // Migrate old fences: move Files into a default tab
            if (fenceInfo.Tabs.Count == 0 && fenceInfo.Files.Count > 0)
            {
                fenceInfo.Tabs.Add(new Model.FenceTab("Main", new System.Collections.Generic.List<string>(fenceInfo.Files)));
                fenceInfo.Files.Clear();
            }
            if (fenceInfo.Tabs.Count == 0)
            {
                fenceInfo.Tabs.Add(new Model.FenceTab("Main"));
            }

            Text = fenceInfo.Name;
            Location = new Point(fenceInfo.PosX, fenceInfo.PosY);

            Width = fenceInfo.Width;
            Height = fenceInfo.Height;

            prevHeight = Height;
            lockedToolStripMenuItem.Checked = fenceInfo.Locked;
            minifyToolStripMenuItem.Checked = fenceInfo.CanMinify;
            Minify();

            FenceManager.Instance.RegisterWindow(this);
            DesktopClickHook.RegisterFenceHandle(Handle);
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

            // Mouse leave (WM_NCMOUSELEAVE — fires when leaving the non-client/title area)
            var myrect = new Rectangle(Location, Size);
            if (m.Msg == 0x02a2 && !myrect.IntersectsWith(new Rectangle(MousePosition, new Size(1, 1))))
            {
                isMouseOver = false;
                hoverTimer.Start();
                Minify();
                selectedItem = null;
                Refresh();
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

                // Right-click on tab bar
                if (pt.Y >= titleHeight && pt.Y < headerHeight)
                {
                    int tabIdx = GetTabIndexAtPoint(pt);
                    if (tabIdx >= 0)
                    {
                        contextMenuTabIndex = tabIdx;
                        tabContextMenu.Show(this, pt);
                    }
                    return;
                }

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
            ActiveFiles.Remove(hoveringItem);
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

        private void FenceWindow_DragOver(object sender, DragEventArgs e)
        {
            var pt = PointToClient(new Point(e.X, e.Y));
            int newDragOverTab = -1;
            if (pt.Y >= titleHeight && pt.Y < headerHeight)
                newDragOverTab = GetTabIndexAtPoint(pt);
            if (newDragOverTab != dragOverTabIndex)
            {
                dragOverTabIndex = newDragOverTab;
                Invalidate();
            }
        }

        private void FenceWindow_DragLeave(object sender, EventArgs e)
        {
            dragOverTabIndex = -1;
            Invalidate();
        }

        private void FenceWindow_DragDrop(object sender, DragEventArgs e)
        {
            var dropped = (string[])e.Data.GetData(DataFormats.FileDrop);
            var clientPt = PointToClient(new Point(e.X, e.Y));
            int targetTab = -1;
            if (clientPt.Y >= titleHeight && clientPt.Y < headerHeight)
                targetTab = GetTabIndexAtPoint(clientPt);

            var targetFiles = (targetTab >= 0) ? fenceInfo.Tabs[targetTab].Files : ActiveFiles;

            foreach (var file in dropped)
            {
                // Check if item is already in another tab (intra-fence move)
                Model.FenceTab sourceTab = null;
                foreach (var t in fenceInfo.Tabs)
                {
                    if (t.Files.Contains(file))
                    {
                        sourceTab = t;
                        break;
                    }
                }

                if (sourceTab != null && targetTab >= 0 && sourceTab != fenceInfo.Tabs[targetTab])
                {
                    sourceTab.Files.Remove(file);
                    fenceInfo.Tabs[targetTab].Files.Add(file);
                }
                else if (sourceTab == null && ItemExists(file))
                {
                    targetFiles.Add(file);
                    HideDesktopIcon(file);
                }
            }
            dragOverTabIndex = -1;
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
                // Check if mousedown is on a tab (for tab reordering)
                if (e.Y >= titleHeight && e.Y < headerHeight)
                {
                    int tabIdx = GetTabIndexAtPoint(e.Location);
                    if (tabIdx >= 0)
                    {
                        draggingTabIndex = tabIdx;
                        tabDragStartPoint = e.Location;
                        tabWasDragged = false;
                        return;
                    }
                }

                dragStartPoint = e.Location;
                dragOutItem = hoveringItem;
            }
        }

        private void FenceWindow_MouseMove(object sender, MouseEventArgs e)
        {
            // Tab drag reordering
            if (e.Button == MouseButtons.Left && draggingTabIndex >= 0)
            {
                if (Math.Abs(e.X - tabDragStartPoint.X) > SystemInformation.DragSize.Width)
                {
                    int targetIdx = GetTabIndexAtPoint(e.Location);
                    if (targetIdx >= 0 && targetIdx != draggingTabIndex)
                    {
                        var tab = fenceInfo.Tabs[draggingTabIndex];
                        fenceInfo.Tabs.RemoveAt(draggingTabIndex);
                        fenceInfo.Tabs.Insert(targetIdx, tab);
                        if (activeTabIndex == draggingTabIndex)
                            activeTabIndex = targetIdx;
                        else if (draggingTabIndex < activeTabIndex && targetIdx >= activeTabIndex)
                            activeTabIndex--;
                        else if (draggingTabIndex > activeTabIndex && targetIdx <= activeTabIndex)
                            activeTabIndex++;
                        draggingTabIndex = targetIdx;
                        tabWasDragged = true;
                        Invalidate();
                    }
                }
                return;
            }

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
                        ActiveFiles.Remove(itemPath);
                        Save();
                    }
                    Refresh();
                    return;
                }
            }
            Invalidate();
        }

        private void FenceWindow_MouseUp(object sender, MouseEventArgs e)
        {
            if (draggingTabIndex >= 0)
            {
                if (tabWasDragged)
                    Save();
                draggingTabIndex = -1;
            }
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
                Height = headerHeight;
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
            // Skip tab switch if we just finished a tab drag
            if (tabWasDragged)
            {
                tabWasDragged = false;
                return;
            }

            var mousePos = PointToClient(MousePosition);

            // Check if click is in tab bar area
            if (mousePos.Y >= titleHeight && mousePos.Y < headerHeight)
            {
                int clickedTab = GetTabIndexAtPoint(mousePos);
                if (clickedTab >= 0)
                {
                    activeTabIndex = clickedTab;
                    scrollOffset = 0;
                    Refresh();
                    return;
                }
                if (IsAddButtonAtPoint(mousePos))
                {
                    AddNewTab();
                    return;
                }
                return;
            }

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

            // Derive colors from accent + opacity
            var accent = Color.FromArgb(fenceInfo.AccentColor);
            float opacityScale = Math.Clamp(fenceInfo.Opacity, 0, 100) / 100f;
            int maxBgAlpha = (int)(255 * opacityScale);
            int idleBgAlpha = (int)(maxBgAlpha * 0.2f);
            int deltaBgAlpha = maxBgAlpha - idleBgAlpha;

            // Background — transparent when idle, opaque on hover
            int bgAlpha = (int)(idleBgAlpha + deltaBgAlpha * hoverAlpha);
            using (var bgBrush = new SolidBrush(Color.FromArgb(bgAlpha, Color.Black)))
                g.FillRectangle(bgBrush, ClientRectangle);

            // Hover border glow — accent color
            if (hoverAlpha > 0)
            {
                int glowAlpha = (int)(60 * hoverAlpha);
                using var borderPen = new Pen(Color.FromArgb(glowAlpha, accent.R, accent.G, accent.B), 1);
                g.DrawRectangle(borderPen, 0, 0, Width - 1, Height - 1);
            }

            // Title background, then text on top
            int titleAlpha = (int)((15 + 60 * hoverAlpha) * opacityScale);
            using (var titleBgBrush = new SolidBrush(Color.FromArgb(titleAlpha, Color.Black)))
                g.FillRectangle(titleBgBrush, new RectangleF(0, 0, Width, titleHeight));
            int textAlpha = (int)(60 + 195 * hoverAlpha);
            using (var titleBrush = new SolidBrush(Color.FromArgb(textAlpha, 255, 255, 255)))
            using (var titleFormat = new StringFormat { Alignment = StringAlignment.Center })
                g.DrawString(Text, titleFont, titleBrush, new PointF(Width / 2, titleOffset), titleFormat);

            // Tab bar
            int tabBgAlpha = (int)((10 + 40 * hoverAlpha) * opacityScale);
            using (var tabBgBrush = new SolidBrush(Color.FromArgb(tabBgAlpha, Color.Black)))
                g.FillRectangle(tabBgBrush, new Rectangle(0, titleHeight, Width, scaledTabBarHeight));

            int tabX = 4;
            var tabMousePos = PointToClient(MousePosition);
            hoveringTabIndex = -1;

            for (int i = 0; i < fenceInfo.Tabs.Count; i++)
            {
                var tab = fenceInfo.Tabs[i];
                var tabTextSize = g.MeasureString(tab.Name, tabFont);
                int tabWidth = (int)tabTextSize.Width + tabPadding * 2;

                if (tabX + tabWidth > Width - addButtonWidth - 4)
                    break; // clip tabs that don't fit

                var tabRect = new Rectangle(tabX, titleHeight, tabWidth, scaledTabBarHeight);
                bool isActive = (i == activeTabIndex);
                bool isTabHover = tabRect.Contains(tabMousePos);
                bool isDragTarget = (i == dragOverTabIndex);

                if (isTabHover)
                    hoveringTabIndex = i;

                if (isDragTarget)
                {
                    using var dragBrush = new SolidBrush(Color.FromArgb(60, accent.R, accent.G, accent.B));
                    g.FillRectangle(dragBrush, tabRect);
                }
                else if (isActive)
                {
                    int activeAlpha = (int)(30 + 80 * hoverAlpha);
                    using var activeBrush = new SolidBrush(Color.FromArgb(activeAlpha, accent.R, accent.G, accent.B));
                    g.FillRectangle(activeBrush, tabRect);
                }
                else if (isTabHover)
                {
                    int hoverTabAlpha = (int)(20 + 50 * hoverAlpha);
                    int dr = (int)(accent.R * 0.8), dg = (int)(accent.G * 0.8), db = (int)(accent.B * 0.8);
                    using var hoverBrush = new SolidBrush(Color.FromArgb(hoverTabAlpha, dr, dg, db));
                    g.FillRectangle(hoverBrush, tabRect);
                }

                if (isActive)
                {
                    using var indicatorPen = new Pen(Color.FromArgb((int)(120 + 135 * hoverAlpha), accent.R, accent.G, accent.B), 2);
                    g.DrawLine(indicatorPen, tabX, titleHeight + scaledTabBarHeight - 1,
                               tabX + tabWidth, titleHeight + scaledTabBarHeight - 1);
                }

                int tabTextAlpha = isActive ? (int)(100 + 155 * hoverAlpha) : (int)(50 + 150 * hoverAlpha);
                using (var tabTextBrush = new SolidBrush(Color.FromArgb(tabTextAlpha, 255, 255, 255)))
                using (var tabFormat = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                    g.DrawString(tab.Name, tabFont, tabTextBrush, tabRect, tabFormat);

                tabX += tabWidth + tabSpacing;
            }

            // "+" add tab button
            var addRect = new Rectangle(tabX, titleHeight, addButtonWidth, scaledTabBarHeight);
            bool addHover = addRect.Contains(tabMousePos);
            int addAlpha = addHover ? (int)(60 + 100 * hoverAlpha) : (int)(40 + 80 * hoverAlpha);
            using (var addBrush = new SolidBrush(Color.FromArgb(addAlpha, 255, 255, 255)))
            using (var addFormat = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                g.DrawString("+", tabFont, addBrush, addRect, addFormat);

            // Items
            var x = itemPadding;
            var y = itemPadding;
            scrollHeight = 0;
            g.Clip = new Region(new Rectangle(0, headerHeight, Width, Height - headerHeight));
            foreach (var file in ActiveFiles)
            {
                var entry = FenceEntry.FromPath(file);
                if (entry == null)
                    continue;

                RenderEntry(g, entry, x, y + headerHeight - scrollOffset);

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

            scrollHeight -= (ClientRectangle.Height - headerHeight);

            // Scroll bars
            if (scrollHeight > 0)
            {
                var contentHeight = Height - headerHeight;
                var scrollbarHeight = contentHeight - scrollHeight;
                using (var scrollBrush = new SolidBrush(Color.FromArgb(150, Color.Black)))
                    g.FillRectangle(scrollBrush, new Rectangle(Width - 5, headerHeight + scrollOffset, 5, scrollbarHeight));

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
            DesktopClickHook.UnregisterFenceHandle(Handle);
            FenceManager.Instance.UnregisterWindow(this);
            // App stays alive via tray icon even with no fences open
        }

        private int GetTabIndexAtPoint(Point pt)
        {
            int tabX = 4;
            using var g = CreateGraphics();
            for (int i = 0; i < fenceInfo.Tabs.Count; i++)
            {
                var tabTextSize = g.MeasureString(fenceInfo.Tabs[i].Name, tabFont);
                int tabWidth = (int)tabTextSize.Width + tabPadding * 2;
                if (tabX + tabWidth > Width - addButtonWidth - 4)
                    return -1;
                var tabRect = new Rectangle(tabX, titleHeight, tabWidth, scaledTabBarHeight);
                if (tabRect.Contains(pt))
                    return i;
                tabX += tabWidth + tabSpacing;
            }
            return -1;
        }

        private bool IsAddButtonAtPoint(Point pt)
        {
            int tabX = 4;
            using var g = CreateGraphics();
            for (int i = 0; i < fenceInfo.Tabs.Count; i++)
            {
                var tabTextSize = g.MeasureString(fenceInfo.Tabs[i].Name, tabFont);
                tabX += (int)tabTextSize.Width + tabPadding * 2 + tabSpacing;
            }
            var addRect = new Rectangle(tabX, titleHeight, addButtonWidth, scaledTabBarHeight);
            return addRect.Contains(pt);
        }

        private void AddNewTab()
        {
            var dialog = new EditDialog("New Tab");
            dialog.TopMost = true;
            if (dialog.ShowDialog(this) == DialogResult.OK)
            {
                fenceInfo.Tabs.Add(new Model.FenceTab(dialog.NewName));
                activeTabIndex = fenceInfo.Tabs.Count - 1;
                scrollOffset = 0;
                Save();
                Refresh();
            }
        }

        private void TabRename_Click(object sender, EventArgs e)
        {
            var tab = fenceInfo.Tabs[contextMenuTabIndex];
            var dialog = new EditDialog(tab.Name);
            dialog.TopMost = true;
            if (dialog.ShowDialog(this) == DialogResult.OK)
            {
                tab.Name = dialog.NewName;
                Save();
                Refresh();
            }
        }

        private void TabDelete_Click(object sender, EventArgs e)
        {
            if (fenceInfo.Tabs.Count <= 1)
                return;
            fenceInfo.Tabs.RemoveAt(contextMenuTabIndex);
            if (activeTabIndex >= fenceInfo.Tabs.Count)
                activeTabIndex = fenceInfo.Tabs.Count - 1;
            Save();
            Refresh();
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
                    Height = headerHeight;
                }
                Refresh();
                Save();
            }
        }

        private void appearanceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var currentColor = Color.FromArgb(fenceInfo.AccentColor);
            var dialog = new AppearanceDialog(currentColor, fenceInfo.Opacity);
            dialog.TopMost = true;
            if (dialog.ShowDialog(this) == DialogResult.OK)
            {
                fenceInfo.AccentColor = dialog.AccentColor.ToArgb();
                fenceInfo.Opacity = dialog.OpacityValue;
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
        private const int SHCNE_ATTRIBUTES = 0x00000800;
        private const int SHCNE_UPDATEDIR = 0x00001000;
        private const uint SHCNF_PATHW = 0x0005;

        private static void NotifyShell(string path)
        {
            var ptr = Marshal.StringToHGlobalUni(path);
            try
            {
                SHChangeNotify(SHCNE_UPDATEITEM, SHCNF_PATHW, ptr, IntPtr.Zero);
                SHChangeNotify(SHCNE_ATTRIBUTES, SHCNF_PATHW, ptr, IntPtr.Zero);
            }
            finally { Marshal.FreeHGlobal(ptr); }

            var dir = Path.GetDirectoryName(path);
            if (dir != null)
            {
                var dirPtr = Marshal.StringToHGlobalUni(dir);
                try { SHChangeNotify(SHCNE_UPDATEDIR, SHCNF_PATHW, dirPtr, IntPtr.Zero); }
                finally { Marshal.FreeHGlobal(dirPtr); }
            }
        }

        private static void HideDesktopIcon(string path)
        {
            try
            {
                var attrs = File.GetAttributes(path);
                File.SetAttributes(path, attrs | FileAttributes.Hidden | FileAttributes.System);
                NotifyShell(path);
                DesktopUtil.RefreshDesktopIcons();
            }
            catch { }
        }

        private static void ShowDesktopIcon(string path)
        {
            try
            {
                var attrs = File.GetAttributes(path);
                File.SetAttributes(path, attrs & ~FileAttributes.Hidden & ~FileAttributes.System);
                NotifyShell(path);
                DesktopUtil.RefreshDesktopIcons();
            }
            catch { }
        }
    }

}

