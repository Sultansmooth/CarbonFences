using CarbonZones.Model;
using CarbonZones.Util;
using CarbonZones.Win32;
using Peter;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
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
        private int iconDrawSize => fenceInfo.IconSize;
        private int itemWidth => iconDrawSize + 30;
        private int itemHeight => iconDrawSize + itemPadding + textHeight;
        private const int textHeight = 35;
        private const int itemPadding = 15;
        private const float shadowDist = 1.5f;

        private readonly FenceInfo fenceInfo;
        public FenceInfo FenceInfo => fenceInfo;

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

        // Internal drag-out state (replaces DoDragDrop)
        private enum ItemDragState { None, Pending, Dragging }
        private ItemDragState itemDragState;
        private string dragItemPath;
        private Point dragItemStart;

        // Live drop feedback while dragging an item within this fence.
        // dragDropSlotIndex: -1 = none, otherwise insertion index in ActiveFiles (0..Count).
        // dragTargetFolder: path key of a folder entry the cursor is directly over.
        private int dragDropSlotIndex = -1;
        private string dragTargetFolder;

        // Captured during Paint so hit-testing uses the exact rendered layout.
        private readonly List<Rectangle> itemLayoutRects = new List<Rectangle>();

        // Throttled-invalidate state — avoids repainting on every mouse pixel.
        private string lastHoveredItemForInvalidate;
        private int lastHoveredTabForInvalidate = -2;

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
            tabContextMenu.Items.Add("Tab Appearance...", null, TabAppearance_Click);
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

        private const int WM_LBUTTONDOWN = 0x0201;
        private const int WM_RBUTTONUP = 0x0205;
        private const int WM_RBUTTONDOWN = 0x0204;
        private const int WM_NCRBUTTONUP = 0x00A5;
        private const int WM_NCLBUTTONDOWN = 0x00A1;
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
            if (m.Msg == 0x02a2 && !myrect.IntersectsWith(new Rectangle(MousePosition, new Size(1, 1)))
                && itemDragState != ItemDragState.Dragging)
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

            // Dismiss open context menus on any mouse click (needed because
            // ShowWithoutActivation / MA_NOACTIVATE prevents normal auto-close)
            if (m.Msg == WM_LBUTTONDOWN || m.Msg == WM_RBUTTONDOWN
                || m.Msg == WM_NCLBUTTONDOWN)
            {
                if (appContextMenu.Visible) appContextMenu.Close();
                if (tabContextMenu.Visible) tabContextMenu.Close();
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
                    ShowShellContextMenuSafe(hoveringItem, PointToScreen(pt));
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
            DesktopUtil.RefreshDesktopIcons();
            Refresh();
        }

        private void contextMenuStrip1_Opening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            deleteItemToolStripMenuItem.Visible = hoveringItem != null;
            addFolderToolStripMenuItem.Visible = hoveringItem == null;
            newFolderToolStripMenuItem.Visible = hoveringItem == null;
        }

        private void addFolderToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (lockedToolStripMenuItem.Checked) return;
            using var dialog = new FolderBrowserDialog();
            dialog.Description = "Select a folder to add to this zone";
            if (dialog.ShowDialog(this) == DialogResult.OK && !ActiveFiles.Contains(dialog.SelectedPath))
            {
                ActiveFiles.Add(dialog.SelectedPath);
                Save();
                Refresh();
            }
        }

        private void newFolderToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (lockedToolStripMenuItem.Checked) return;

            var dialog = new EditDialog("New Folder");
            dialog.TopMost = true;
            if (dialog.ShowDialog(this) != DialogResult.OK) return;

            var rawName = (dialog.NewName ?? "").Trim();
            if (string.IsNullOrEmpty(rawName)) return;

            // Strip characters Windows forbids in file/folder names.
            var invalid = Path.GetInvalidFileNameChars();
            var sanitized = new string(rawName.Where(c => Array.IndexOf(invalid, c) < 0).ToArray()).Trim();
            if (string.IsNullOrEmpty(sanitized)) return;

            var zoneRoot = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "Carbon Zones",
                SanitizeForPath(fenceInfo.Name));
            var target = Path.Combine(zoneRoot, sanitized);

            // Disambiguate if a folder with that name already exists in this zone root.
            int suffix = 2;
            while (Directory.Exists(target))
            {
                target = Path.Combine(zoneRoot, $"{sanitized} ({suffix++})");
            }

            try
            {
                Directory.CreateDirectory(target);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Could not create folder:\n{ex.Message}", "New Folder",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            ActiveFiles.Add(target);
            Save();
            Refresh();
        }

        private static string SanitizeForPath(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "Zone";
            var invalid = Path.GetInvalidFileNameChars();
            var cleaned = new string(name.Where(c => Array.IndexOf(invalid, c) < 0).ToArray()).Trim();
            return string.IsNullOrEmpty(cleaned) ? "Zone" : cleaned;
        }

        private void FenceWindow_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop) && !lockedToolStripMenuItem.Checked)
                e.Effect = DragDropEffects.Copy;
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
            DesktopUtil.RefreshDesktopIcons();
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

                // Start potential item drag (internal capture, no DoDragDrop)
                if (hoveringItem != null && !lockedToolStripMenuItem.Checked)
                {
                    itemDragState = ItemDragState.Pending;
                    dragItemPath = hoveringItem;
                    dragItemStart = e.Location;
                }
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

            // Internal item drag — mouse capture based (no DoDragDrop)
            if (e.Button == MouseButtons.Left && itemDragState == ItemDragState.Pending)
            {
                if (Math.Abs(e.X - dragItemStart.X) > SystemInformation.DragSize.Width ||
                    Math.Abs(e.Y - dragItemStart.Y) > SystemInformation.DragSize.Height)
                {
                    itemDragState = ItemDragState.Dragging;
                    Capture = true;
                    Cursor = Cursors.Hand;
                }
            }

            // While dragging an item, compute live drop feedback (folder target or
            // insertion slot). UpdateDragDropFeedback only invalidates on change.
            if (itemDragState == ItemDragState.Dragging && dragItemPath != null)
            {
                UpdateDragDropFeedback(e.Location);
                return;
            }

            // Throttled hover invalidation: a full repaint per mouse pixel (with
            // shell icon extraction inside Paint) was the cause of "low-fps"
            // chop. We hit-test against the last-painted layout and only ask
            // for a redraw when the hovered item or tab actually changes.
            string newHover = HitTestItem(e.Location);
            int newTabHover = (e.Y >= titleHeight && e.Y < headerHeight)
                ? GetTabIndexAtPoint(e.Location)
                : -1;

            if (newHover != lastHoveredItemForInvalidate || newTabHover != lastHoveredTabForInvalidate)
            {
                lastHoveredItemForInvalidate = newHover;
                lastHoveredTabForInvalidate = newTabHover;
                Invalidate();
            }
        }

        private string HitTestItem(Point p)
        {
            // itemLayoutRects is populated during Paint; for indices in range,
            // each entry corresponds 1:1 with ActiveFiles[i].
            int n = Math.Min(itemLayoutRects.Count, ActiveFiles.Count);
            for (int i = 0; i < n; i++)
            {
                if (itemLayoutRects[i].Contains(p))
                    return ActiveFiles[i];
            }
            return null;
        }

        private void UpdateDragDropFeedback(Point clientPos)
        {
            string newFolder = null;
            int newSlot = -1;

            bool inContentArea = clientPos.Y >= headerHeight && ClientRectangle.Contains(clientPos);
            if (inContentArea)
            {
                // Folder highlight takes priority: if the cursor is over a folder entry
                // that isn't the one being dragged, that's a "move into folder" target.
                if (!string.IsNullOrEmpty(hoveringItem) &&
                    hoveringItem != dragItemPath &&
                    Directory.Exists(FileStaging.GetEffectivePath(hoveringItem)))
                {
                    newFolder = hoveringItem;
                }
                else
                {
                    newSlot = GetInsertSlotAtPoint(clientPos);
                }
            }

            if (newFolder != dragTargetFolder || newSlot != dragDropSlotIndex)
            {
                dragTargetFolder = newFolder;
                dragDropSlotIndex = newSlot;
                Invalidate();
            }
        }

        /// <summary>
        /// Returns the insertion index within ActiveFiles that corresponds to the
        /// given client-space point. Uses the rectangles captured during the last Paint.
        /// Returns ActiveFiles.Count when the point is past the end.
        /// </summary>
        private int GetInsertSlotAtPoint(Point p)
        {
            if (itemLayoutRects.Count == 0)
                return 0;

            // Find the closest cell center, then decide "before" or "after" by X.
            int bestIdx = -1;
            double bestDist = double.MaxValue;
            for (int i = 0; i < itemLayoutRects.Count; i++)
            {
                var r = itemLayoutRects[i];
                double dx = p.X - (r.X + r.Width / 2.0);
                double dy = p.Y - (r.Y + r.Height / 2.0);
                double d = dx * dx + dy * dy;
                if (d < bestDist) { bestDist = d; bestIdx = i; }
            }

            if (bestIdx < 0) return itemLayoutRects.Count;

            var best = itemLayoutRects[bestIdx];
            var centerX = best.X + best.Width / 2;
            return (p.X < centerX) ? bestIdx : bestIdx + 1;
        }

        private void FenceWindow_MouseUp(object sender, MouseEventArgs e)
        {
            // Tab drag reorder finish
            if (draggingTabIndex >= 0)
            {
                if (tabWasDragged)
                    Save();
                draggingTabIndex = -1;
            }

            // Item drag resolution
            if (itemDragState == ItemDragState.Dragging && dragItemPath != null)
            {
                // Save drag info and reset state BEFORE releasing capture,
                // because Capture=false triggers OnMouseCaptureChanged which
                // would otherwise clear the state before we can resolve the drop.
                var droppedPath = dragItemPath;
                itemDragState = ItemDragState.None;
                dragItemPath = null;
                dragDropSlotIndex = -1;
                dragTargetFolder = null;

                Capture = false;
                Cursor = Cursors.Default;

                var screenPos = Control.MousePosition;
                var clientPos = PointToClient(screenPos);

                // Check if dropped on another fence
                var targetFence = FenceManager.Instance.GetFenceAtScreenPoint(screenPos);

                if (targetFence != null && targetFence != this)
                {
                    // Inter-fence transfer
                    targetFence.AcceptItem(droppedPath);
                    ActiveFiles.Remove(droppedPath);
                    Save();
                }
                else if (ClientRectangle.Contains(clientPos))
                {
                    // Dropped back on same fence — check if on a different tab
                    if (clientPos.Y >= titleHeight && clientPos.Y < headerHeight)
                    {
                        int targetTab = GetTabIndexAtPoint(clientPos);
                        if (targetTab >= 0 && targetTab != activeTabIndex)
                        {
                            ActiveFiles.Remove(droppedPath);
                            fenceInfo.Tabs[targetTab].Files.Add(droppedPath);
                            Save();
                        }
                    }
                    else if (!string.IsNullOrEmpty(hoveringItem) &&
                             hoveringItem != droppedPath &&
                             Directory.Exists(FileStaging.GetEffectivePath(hoveringItem)))
                    {
                        // Dropped onto a folder entry — move the source into that folder on disk.
                        // Use effective paths so staged items (which live under __staged) work.
                        if (TryMoveItemIntoFolder(droppedPath, hoveringItem))
                        {
                            ActiveFiles.Remove(droppedPath);
                            Save();
                        }
                    }
                    else
                    {
                        // Dropped within content area but not on a folder — reorder.
                        int srcIdx = ActiveFiles.IndexOf(droppedPath);
                        int dstSlot = GetInsertSlotAtPoint(clientPos);
                        if (srcIdx >= 0 && dstSlot != srcIdx && dstSlot != srcIdx + 1)
                        {
                            ActiveFiles.RemoveAt(srcIdx);
                            if (dstSlot > srcIdx) dstSlot--;
                            if (dstSlot < 0) dstSlot = 0;
                            if (dstSlot > ActiveFiles.Count) dstSlot = ActiveFiles.Count;
                            ActiveFiles.Insert(dstSlot, droppedPath);
                            Save();
                        }
                        // else: no net change — item stays put.
                    }
                }
                else
                {
                    // Dropped on desktop (or outside any fence) — unhide and remove
                    ShowDesktopIcon(droppedPath);
                    ActiveFiles.Remove(droppedPath);
                    Save();
                    DesktopUtil.RefreshDesktopIcons();
                }

                Refresh();
                return;
            }

            // Pending drag that never crossed threshold — just reset
            if (itemDragState == ItemDragState.Pending)
            {
                itemDragState = ItemDragState.None;
                dragItemPath = null;
            }
        }

        protected override void OnMouseCaptureChanged(EventArgs e)
        {
            if (itemDragState == ItemDragState.Dragging)
            {
                // Capture lost unexpectedly (Alt+Tab, another app, etc.) — cancel drag
                itemDragState = ItemDragState.None;
                dragItemPath = null;
                dragDropSlotIndex = -1;
                dragTargetFolder = null;
                Cursor = Cursors.Default;
                Refresh();
            }
            base.OnMouseCaptureChanged(e);
        }

        /// <summary>
        /// Moves a fence entry (file or folder) into the given target folder on disk.
        /// Handles the case where either path is currently staged, and resolves name
        /// collisions by appending " (2)", " (3)", etc.
        /// </summary>
        private bool TryMoveItemIntoFolder(string sourceKey, string targetFolderKey)
        {
            try
            {
                var src = FileStaging.GetEffectivePath(sourceKey);
                var destFolder = FileStaging.GetEffectivePath(targetFolderKey);
                if (!Directory.Exists(destFolder)) return false;

                bool srcIsDir = Directory.Exists(src);
                bool srcIsFile = File.Exists(src);
                if (!srcIsDir && !srcIsFile) return false;

                var baseName = Path.GetFileName(src.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
                var dest = Path.Combine(destFolder, baseName);

                int n = 2;
                while (File.Exists(dest) || Directory.Exists(dest))
                {
                    var stem = Path.GetFileNameWithoutExtension(baseName);
                    var ext = Path.GetExtension(baseName);
                    dest = Path.Combine(destFolder, $"{stem} ({n++}){ext}");
                }

                if (srcIsDir)
                    Directory.Move(src, dest);
                else
                    File.Move(src, dest);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to move item into folder: {ex}");
                return false;
            }
        }

        public void AcceptItem(string filePath)
        {
            if (lockedToolStripMenuItem.Checked) return;
            ActiveFiles.Add(filePath);
            Save();
            Refresh();
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
            if (itemDragState == ItemDragState.Dragging) return;
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

            // Derive colors — active tab overrides fence defaults
            var activeTab = fenceInfo.Tabs[activeTabIndex];
            var accent = activeTab.AccentColor != 0 ? Color.FromArgb(activeTab.AccentColor) : Color.FromArgb(fenceInfo.AccentColor);
            var boxBase = activeTab.BoxColor != 0 ? Color.FromArgb(activeTab.BoxColor) : (fenceInfo.BoxColor != 0 ? Color.FromArgb(fenceInfo.BoxColor) : Color.Black);
            var labelBase = activeTab.LabelColor != 0 ? Color.FromArgb(activeTab.LabelColor) : (fenceInfo.LabelColor != 0 ? Color.FromArgb(fenceInfo.LabelColor) : Color.Black);
            float opacityScale = Math.Clamp(fenceInfo.Opacity, 0, 100) / 100f;
            int maxBgAlpha = (int)(255 * opacityScale);
            int idleBgAlpha = (int)(maxBgAlpha * 0.2f);
            int deltaBgAlpha = maxBgAlpha - idleBgAlpha;

            // Background — transparent when idle, opaque on hover
            int bgAlpha = (int)(idleBgAlpha + deltaBgAlpha * hoverAlpha);
            using (var bgBrush = new SolidBrush(Color.FromArgb(bgAlpha, boxBase.R, boxBase.G, boxBase.B)))
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
            using (var titleBgBrush = new SolidBrush(Color.FromArgb(titleAlpha, labelBase.R, labelBase.G, labelBase.B)))
                g.FillRectangle(titleBgBrush, new RectangleF(0, 0, Width, titleHeight));
            int textAlpha = (int)(60 + 195 * hoverAlpha);
            using (var titleBrush = new SolidBrush(Color.FromArgb(textAlpha, 255, 255, 255)))
            using (var titleFormat = new StringFormat { Alignment = StringAlignment.Center })
                g.DrawString(Text, titleFont, titleBrush, new PointF(Width / 2, titleOffset), titleFormat);

            // Tab bar
            int tabBgAlpha = (int)((10 + 40 * hoverAlpha) * opacityScale);
            using (var tabBgBrush = new SolidBrush(Color.FromArgb(tabBgAlpha, labelBase.R, labelBase.G, labelBase.B)))
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

                // Per-tab accent color (falls back to fence accent)
                var tabAccent = tab.AccentColor != 0 ? Color.FromArgb(tab.AccentColor) : Color.FromArgb(fenceInfo.AccentColor);

                if (isDragTarget)
                {
                    using var dragBrush = new SolidBrush(Color.FromArgb(60, tabAccent.R, tabAccent.G, tabAccent.B));
                    g.FillRectangle(dragBrush, tabRect);
                }
                else if (isActive)
                {
                    int activeAlpha = (int)(30 + 80 * hoverAlpha);
                    using var activeBrush = new SolidBrush(Color.FromArgb(activeAlpha, tabAccent.R, tabAccent.G, tabAccent.B));
                    g.FillRectangle(activeBrush, tabRect);
                }
                else if (isTabHover)
                {
                    int hoverTabAlpha = (int)(20 + 50 * hoverAlpha);
                    int dr = (int)(tabAccent.R * 0.8), dg = (int)(tabAccent.G * 0.8), db = (int)(tabAccent.B * 0.8);
                    using var hoverBrush = new SolidBrush(Color.FromArgb(hoverTabAlpha, dr, dg, db));
                    g.FillRectangle(hoverBrush, tabRect);
                }

                if (isActive)
                {
                    using var indicatorPen = new Pen(Color.FromArgb((int)(120 + 135 * hoverAlpha), tabAccent.R, tabAccent.G, tabAccent.B), 2);
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
            itemLayoutRects.Clear();
            g.Clip = new Region(new Rectangle(0, headerHeight, Width, Height - headerHeight));
            for (int idx = 0; idx < ActiveFiles.Count; idx++)
            {
                var file = ActiveFiles[idx];
                var effectivePath = FileStaging.GetEffectivePath(file);
                var entry = FenceEntry.FromPath(effectivePath);

                // Always push a rect so itemLayoutRects indices stay aligned with ActiveFiles.
                var cellTop = y + headerHeight - scrollOffset;
                itemLayoutRects.Add(new Rectangle(x, cellTop, itemWidth, itemHeight));

                if (entry != null)
                {
                    RenderEntry(g, entry, file, x, cellTop);
                }

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

            // Drop-slot indicator while dragging within the same tab.
            if (itemDragState == ItemDragState.Dragging && dragDropSlotIndex >= 0 && dragTargetFolder == null)
            {
                DrawDropSlotIndicator(g, accent);
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

            // Sync throttled-invalidate state with what was just rendered, so
            // MouseMove only triggers a repaint when the hover actually changes.
            lastHoveredItemForInvalidate = hoveringItem;
            lastHoveredTabForInvalidate = hoveringTabIndex;
        }

        private void DrawDropSlotIndicator(Graphics g, Color accent)
        {
            // Derive the position of the insertion slot from the layout rects.
            // If the slot index equals Count, draw after the last item (same row if space, else new row).
            int slot = dragDropSlotIndex;
            if (itemLayoutRects.Count == 0)
                return;

            Rectangle refRect;
            bool drawBefore;
            if (slot <= 0)
            {
                refRect = itemLayoutRects[0];
                drawBefore = true;
            }
            else if (slot >= itemLayoutRects.Count)
            {
                refRect = itemLayoutRects[itemLayoutRects.Count - 1];
                drawBefore = false;
            }
            else
            {
                // Prefer the right edge of slot-1 when it's on the same row as slot,
                // otherwise draw on the left edge of slot (to indicate wrap).
                var prev = itemLayoutRects[slot - 1];
                var next = itemLayoutRects[slot];
                if (prev.Y == next.Y)
                {
                    refRect = prev;
                    drawBefore = false;
                }
                else
                {
                    refRect = next;
                    drawBefore = true;
                }
            }

            int lineX = drawBefore ? refRect.Left - itemPadding / 2 : refRect.Right + itemPadding / 2;
            int top = refRect.Top - 2;
            int bottom = refRect.Bottom + 2;

            using var pen = new Pen(Color.FromArgb(220, accent.R, accent.G, accent.B), 2f);
            g.DrawLine(pen, lineX, top, lineX, bottom);
            // Small end caps to make the indicator more visible.
            g.DrawLine(pen, lineX - 3, top, lineX + 3, top);
            g.DrawLine(pen, lineX - 3, bottom, lineX + 3, bottom);
        }

        private void RenderEntry(Graphics g, FenceEntry entry, string originalPath, int x, int y)
        {
            var largeBmp = entry.ExtractLargeIcon(thumbnailProvider);
            var name = entry.Name;

            var textPosition = new PointF(x, y + iconDrawSize + 5);
            var textMaxSize = new SizeF(itemWidth, textHeight);

            using var stringFormat = new StringFormat { Alignment = StringAlignment.Center, Trimming = StringTrimming.EllipsisCharacter };

            var textSize = g.MeasureString(name, iconFont, textMaxSize, stringFormat);
            var outlineRect = new Rectangle(x - 2, y - 2, itemWidth + 2, iconDrawSize + (int)textSize.Height + 5 + 2);
            var outlineRectInner = outlineRect.Shrink(1);

            var mousePos = PointToClient(MousePosition);
            var mouseOver = mousePos.X >= x && mousePos.Y >= y && mousePos.X < x + outlineRect.Width && mousePos.Y < y + outlineRect.Height;

            if (mouseOver)
            {
                hoveringItem = originalPath;
                hasHoverUpdated = true;
            }

            if (mouseOver && shouldUpdateSelection)
            {
                selectedItem = originalPath;
                shouldUpdateSelection = false;
                hasSelectionUpdated = true;
            }

            if (mouseOver && shouldRunDoubleClick)
            {
                shouldRunDoubleClick = false;
                entry.Open();
            }

            if (selectedItem == originalPath)
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

            // Folder drop target highlight — drawn during an internal item drag when
            // the cursor sits on a folder entry other than the dragged item.
            if (dragTargetFolder != null && originalPath == dragTargetFolder &&
                itemDragState == ItemDragState.Dragging)
            {
                var tabForAccent = fenceInfo.Tabs[activeTabIndex];
                var accent = tabForAccent.AccentColor != 0
                    ? Color.FromArgb(tabForAccent.AccentColor)
                    : Color.FromArgb(fenceInfo.AccentColor);
                using var dropFill = new SolidBrush(Color.FromArgb(80, accent.R, accent.G, accent.B));
                using var dropBorder = new Pen(Color.FromArgb(220, accent.R, accent.G, accent.B), 2f);
                g.FillRectangle(dropFill, outlineRect);
                g.DrawRectangle(dropBorder, outlineRectInner);
            }

            int iconX = x + itemWidth / 2 - iconDrawSize / 2;
            if (largeBmp != null)
            {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.DrawImage(largeBmp,
                    new Rectangle(iconX, y, iconDrawSize, iconDrawSize),
                    new Rectangle(0, 0, largeBmp.Width, largeBmp.Height),
                    GraphicsUnit.Pixel);
            }
            else
            {
                var icon = entry.ExtractIcon(thumbnailProvider);
                using (var bmp = icon.ToBitmap())
                {
                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    g.DrawImage(bmp,
                        new Rectangle(iconX, y, iconDrawSize, iconDrawSize),
                        new Rectangle(0, 0, bmp.Width, bmp.Height),
                        GraphicsUnit.Pixel);
                }
            }
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

        private void TabAppearance_Click(object sender, EventArgs e)
        {
            var tab = fenceInfo.Tabs[contextMenuTabIndex];
            var currentAccent = tab.AccentColor != 0 ? Color.FromArgb(tab.AccentColor) : Color.FromArgb(fenceInfo.AccentColor);
            var currentLabel = tab.LabelColor != 0 ? Color.FromArgb(tab.LabelColor) : (fenceInfo.LabelColor != 0 ? Color.FromArgb(fenceInfo.LabelColor) : Color.Black);
            var currentBox = tab.BoxColor != 0 ? Color.FromArgb(tab.BoxColor) : (fenceInfo.BoxColor != 0 ? Color.FromArgb(fenceInfo.BoxColor) : Color.Black);
            var dialog = new AppearanceDialog(currentAccent, fenceInfo.Opacity, currentLabel, currentBox);
            dialog.Text = $"Tab Appearance — {tab.Name}";
            dialog.TopMost = true;
            if (dialog.ShowDialog(this) == DialogResult.OK)
            {
                tab.AccentColor = (dialog.AccentColor.ToArgb() == fenceInfo.AccentColor) ? 0 : dialog.AccentColor.ToArgb();
                tab.LabelColor = (dialog.LabelColor.R == 0 && dialog.LabelColor.G == 0 && dialog.LabelColor.B == 0) ? 0 : dialog.LabelColor.ToArgb();
                tab.BoxColor = (dialog.BoxColor.R == 0 && dialog.BoxColor.G == 0 && dialog.BoxColor.B == 0) ? 0 : dialog.BoxColor.ToArgb();
                Refresh();
                Save();
            }
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
            var currentAccent = Color.FromArgb(fenceInfo.AccentColor);
            var currentLabel = fenceInfo.LabelColor != 0 ? Color.FromArgb(fenceInfo.LabelColor) : Color.Black;
            var currentBox = fenceInfo.BoxColor != 0 ? Color.FromArgb(fenceInfo.BoxColor) : Color.Black;
            var dialog = new AppearanceDialog(currentAccent, fenceInfo.Opacity, currentLabel, currentBox);
            dialog.TopMost = true;
            if (dialog.ShowDialog(this) == DialogResult.OK)
            {
                fenceInfo.AccentColor = dialog.AccentColor.ToArgb();
                fenceInfo.Opacity = dialog.OpacityValue;
                fenceInfo.LabelColor = (dialog.LabelColor.R == 0 && dialog.LabelColor.G == 0 && dialog.LabelColor.B == 0) ? 0 : dialog.LabelColor.ToArgb();
                fenceInfo.BoxColor = (dialog.BoxColor.R == 0 && dialog.BoxColor.G == 0 && dialog.BoxColor.B == 0) ? 0 : dialog.BoxColor.ToArgb();
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
                ShowShellContextMenuSafe(hoveringItem, MousePosition);
            }
            else
            {
                appContextMenu.Show(this, e.Location);
            }
        }

        private void FenceWindow_MouseWheel(object sender, MouseEventArgs e)
        {
            // Ctrl+Scroll = resize icons
            if (ModifierKeys.HasFlag(Keys.Control))
            {
                if (e is HandledMouseEventArgs hme) hme.Handled = true;
                int step = Math.Sign(e.Delta) * 10;
                fenceInfo.IconSize = Math.Clamp(fenceInfo.IconSize + step, 32, 96);
                Save();
                Refresh();
                return;
            }

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

        /// <summary>
        /// Shows the shell context menu for a fenced item, then checks if the
        /// file was renamed or deleted via the menu. Updates the Files list
        /// accordingly so items don't silently vanish.
        /// </summary>
        private void ShowShellContextMenuSafe(string originalPath, Point screenPos)
        {
            var effectivePath = FileStaging.GetEffectivePath(originalPath);
            // Use the correct overload for files vs directories
            if (Directory.Exists(effectivePath))
                shellContextMenu.ShowContextMenu(new[] { new DirectoryInfo(effectivePath) }, screenPos);
            else
                shellContextMenu.ShowContextMenu(new[] { new FileInfo(effectivePath) }, screenPos);
            // Don't try to detect renames here — Properties dialogs rename
            // asynchronously after the context menu closes. The staging
            // FileSystemWatcher reconciliation handles renames and deletes.
            Refresh();
        }

        private static void HideDesktopIcon(string path)
        {
            FileStaging.Stage(path);
        }

        private static void ShowDesktopIcon(string path)
        {
            FileStaging.Unstage(path);
        }
    }

}

