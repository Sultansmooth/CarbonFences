using System.Drawing;
using System.Windows.Forms;

namespace CarbonZones
{
    public class SelectionOverlay : Form
    {
        private Rectangle selectionRect;

        public SelectionOverlay()
        {
            FormBorderStyle = FormBorderStyle.None;
            BackColor = Color.Fuchsia;
            TransparencyKey = Color.Fuchsia;
            TopMost = true;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            Bounds = SystemInformation.VirtualScreen;
            DoubleBuffered = true;
        }

        protected override bool ShowWithoutActivation => true;

        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ExStyle |= 0x00000080; // WS_EX_TOOLWINDOW
                cp.ExStyle |= 0x08000000; // WS_EX_NOACTIVATE
                cp.ExStyle |= 0x00000020; // WS_EX_TRANSPARENT (click-through)
                return cp;
            }
        }

        public void UpdateSelection(Rectangle rect)
        {
            selectionRect = rect;
            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            if (selectionRect.Width > 0 && selectionRect.Height > 0)
            {
                // Convert screen coords to client coords
                var localRect = new Rectangle(
                    selectionRect.X - Left,
                    selectionRect.Y - Top,
                    selectionRect.Width,
                    selectionRect.Height);
                using var pen = new Pen(Color.FromArgb(0, 120, 215), 2);
                using var brush = new SolidBrush(Color.FromArgb(180, 210, 235));
                e.Graphics.FillRectangle(brush, localRect);
                e.Graphics.DrawRectangle(pen, localRect);
            }
        }
    }
}
