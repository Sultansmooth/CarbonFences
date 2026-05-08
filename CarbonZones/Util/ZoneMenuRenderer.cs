using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace CarbonZones.Util
{
    /// <summary>
    /// Custom ToolStrip renderer that styles right-click menus to match the
    /// dark/translucent fence aesthetic. Uses a delegate so the accent color
    /// stays in sync with the fence's current Appearance settings.
    /// </summary>
    public class ZoneMenuRenderer : ToolStripProfessionalRenderer
    {
        private readonly Func<Color> getAccent;

        // Background tone — near-black with a hint of cool. Opaque so text
        // legibility is unconditional regardless of what's behind the menu.
        private static readonly Color BgTop = Color.FromArgb(34, 34, 40);
        private static readonly Color BgBottom = Color.FromArgb(22, 22, 28);

        private const int CornerRadius = 8;

        public ZoneMenuRenderer(Func<Color> getAccent)
            : base(new ZoneColorTable(getAccent))
        {
            this.getAccent = getAccent ?? (() => Color.FromArgb(100, 160, 230));
            RoundedEdges = false; // we draw our own corners
        }

        protected override void OnRenderToolStripBackground(ToolStripRenderEventArgs e)
        {
            var g = e.Graphics;
            var bounds = new Rectangle(Point.Empty, e.ToolStrip.Size);
            g.SmoothingMode = SmoothingMode.AntiAlias;

            using var path = BuildRoundedPath(bounds, CornerRadius);
            using var brush = new LinearGradientBrush(bounds, BgTop, BgBottom, LinearGradientMode.Vertical);
            g.FillPath(brush, path);
        }

        protected override void OnRenderToolStripBorder(ToolStripRenderEventArgs e)
        {
            var g = e.Graphics;
            var bounds = new Rectangle(0, 0, e.ToolStrip.Width - 1, e.ToolStrip.Height - 1);
            var accent = getAccent();
            g.SmoothingMode = SmoothingMode.AntiAlias;

            using var path = BuildRoundedPath(bounds, CornerRadius);
            // Inner subtle white-ish highlight for a "glass" feel.
            using var innerPen = new Pen(Color.FromArgb(28, 255, 255, 255), 1f);
            g.DrawPath(innerPen, path);

            // Outer accent-colored stroke for color identity.
            using var accentPen = new Pen(Color.FromArgb(140, accent.R, accent.G, accent.B), 1f);
            g.DrawPath(accentPen, path);

            // Apply the rounded region to the window itself so the corners
            // don't show the underlying rectangular OS frame.
            try
            {
                using var regionPath = BuildRoundedPath(new Rectangle(0, 0, e.ToolStrip.Width, e.ToolStrip.Height), CornerRadius);
                e.ToolStrip.Region = new Region(regionPath);
            }
            catch { }
        }

        protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e)
        {
            var item = e.Item;
            if (!item.Selected && !(item is ToolStripMenuItem mi && mi.DropDown.Visible))
                return;

            var g = e.Graphics;
            var accent = getAccent();
            var rect = new Rectangle(3, 1, item.Width - 6, item.Height - 2);
            int alpha = item.Pressed ? 200 : 150;

            g.SmoothingMode = SmoothingMode.AntiAlias;
            using var path = BuildRoundedPath(rect, 4);
            using var fillBrush = new SolidBrush(Color.FromArgb(alpha, accent.R, accent.G, accent.B));
            g.FillPath(fillBrush, path);

            using var borderPen = new Pen(Color.FromArgb(220, accent.R, accent.G, accent.B), 1f);
            g.DrawPath(borderPen, path);
        }

        protected override void OnRenderItemText(ToolStripItemTextRenderEventArgs e)
        {
            // White text everywhere; dim slightly when disabled.
            e.TextColor = e.Item.Enabled
                ? Color.FromArgb(245, 245, 247)
                : Color.FromArgb(120, 245, 245, 247);
            base.OnRenderItemText(e);
        }

        protected override void OnRenderArrow(ToolStripArrowRenderEventArgs e)
        {
            e.ArrowColor = Color.FromArgb(220, 245, 245, 247);
            base.OnRenderArrow(e);
        }

        protected override void OnRenderSeparator(ToolStripSeparatorRenderEventArgs e)
        {
            var g = e.Graphics;
            var b = e.Item.Bounds;
            int y = b.Height / 2;
            using var pen = new Pen(Color.FromArgb(40, 255, 255, 255), 1f);
            g.DrawLine(pen, 8, y, b.Width - 8, y);
        }

        protected override void OnRenderImageMargin(ToolStripRenderEventArgs e)
        {
            // Suppress the default lighter gutter — let the gradient background show through.
        }

        protected override void OnRenderItemCheck(ToolStripItemImageRenderEventArgs e)
        {
            // Render a clean accent-colored checkmark (matches fence accent).
            var g = e.Graphics;
            var accent = getAccent();
            var r = e.ImageRectangle;
            if (r.Width <= 0 || r.Height <= 0) return;

            g.SmoothingMode = SmoothingMode.AntiAlias;
            using var pen = new Pen(accent, 2f) { StartCap = LineCap.Round, EndCap = LineCap.Round };
            int pad = Math.Max(2, r.Width / 5);
            var p1 = new Point(r.Left + pad, r.Top + r.Height / 2);
            var p2 = new Point(r.Left + r.Width / 2 - 1, r.Bottom - pad);
            var p3 = new Point(r.Right - pad, r.Top + pad);
            g.DrawLines(pen, new[] { p1, p2, p3 });
        }

        private static GraphicsPath BuildRoundedPath(Rectangle r, int radius)
        {
            int d = radius * 2;
            if (d > r.Width) d = r.Width;
            if (d > r.Height) d = r.Height;

            var path = new GraphicsPath();
            if (d <= 0)
            {
                path.AddRectangle(r);
                return path;
            }

            path.AddArc(r.Left, r.Top, d, d, 180, 90);
            path.AddArc(r.Right - d, r.Top, d, d, 270, 90);
            path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            path.AddArc(r.Left, r.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }

        private class ZoneColorTable : ProfessionalColorTable
        {
            private readonly Func<Color> getAccent;
            public ZoneColorTable(Func<Color> getAccent) { this.getAccent = getAccent; }

            // Most paint paths are taken over by ZoneMenuRenderer; this table
            // exists to neutralize defaults that would otherwise show through
            // (gradients on the image margin, separator highlights, etc.).
            public override Color ToolStripDropDownBackground => Color.FromArgb(28, 28, 34);
            public override Color MenuBorder => Color.FromArgb(140, getAccent());
            public override Color MenuItemBorder => Color.Transparent;
            public override Color MenuItemSelected => Color.FromArgb(150, getAccent());
            public override Color MenuItemSelectedGradientBegin => Color.FromArgb(150, getAccent());
            public override Color MenuItemSelectedGradientEnd => Color.FromArgb(150, getAccent());
            public override Color MenuItemPressedGradientBegin => Color.FromArgb(200, getAccent());
            public override Color MenuItemPressedGradientEnd => Color.FromArgb(200, getAccent());
            public override Color ImageMarginGradientBegin => Color.Transparent;
            public override Color ImageMarginGradientMiddle => Color.Transparent;
            public override Color ImageMarginGradientEnd => Color.Transparent;
            public override Color SeparatorDark => Color.FromArgb(40, 255, 255, 255);
            public override Color SeparatorLight => Color.FromArgb(20, 255, 255, 255);
            public override Color CheckBackground => Color.Transparent;
            public override Color CheckPressedBackground => Color.Transparent;
            public override Color CheckSelectedBackground => Color.Transparent;
            public override Color ButtonSelectedBorder => Color.Transparent;
        }
    }
}
