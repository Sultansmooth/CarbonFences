using System;
using System.Drawing;
using System.Windows.Forms;

namespace CarbonZones
{
    public class AppearanceDialog : Form
    {
        private static readonly Color[] Presets = new[]
        {
            Color.FromArgb(100, 160, 230), // Blue (default)
            Color.FromArgb(140, 100, 230), // Purple
            Color.FromArgb(60, 180, 180),  // Teal
            Color.FromArgb(80, 180, 100),  // Green
            Color.FromArgb(220, 80, 80),   // Red
            Color.FromArgb(230, 150, 60),  // Orange
            Color.FromArgb(220, 100, 160), // Pink
            Color.FromArgb(60, 180, 220),  // Cyan
            Color.FromArgb(210, 200, 60),  // Yellow
            Color.FromArgb(140, 140, 140), // Gray
        };

        private Color selectedColor;
        private readonly Panel[] swatchPanels;
        private readonly TrackBar opacityTrack;
        private readonly Label opacityLabel;

        public Color AccentColor => selectedColor;
        public int OpacityValue => opacityTrack.Value;

        public AppearanceDialog(Color currentColor, int currentOpacity)
        {
            selectedColor = currentColor;

            Text = "Appearance";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            Size = new Size(310, 250);
            BackColor = Color.FromArgb(32, 32, 32);
            ForeColor = Color.White;

            // Color label
            var colorLabel = new Label
            {
                Text = "Accent Color",
                Location = new Point(12, 12),
                AutoSize = true,
                ForeColor = Color.FromArgb(200, 200, 200)
            };
            Controls.Add(colorLabel);

            // Color swatches — 5 per row, 2 rows
            swatchPanels = new Panel[Presets.Length];
            for (int i = 0; i < Presets.Length; i++)
            {
                var swatch = new Panel
                {
                    Size = new Size(28, 28),
                    Location = new Point(12 + (i % 5) * 34, 34 + (i / 5) * 34),
                    BackColor = Presets[i],
                    Cursor = Cursors.Hand
                };
                int idx = i;
                swatch.Click += (s, e) => SelectColor(Presets[idx]);
                swatch.Paint += SwatchPaint;
                swatchPanels[i] = swatch;
                Controls.Add(swatch);
            }

            // Custom color button
            var customBtn = new Button
            {
                Text = "Custom...",
                Location = new Point(190, 34),
                Size = new Size(90, 28),
                FlatStyle = FlatStyle.Flat,
                ForeColor = Color.White,
                BackColor = Color.FromArgb(55, 55, 55)
            };
            customBtn.FlatAppearance.BorderColor = Color.FromArgb(80, 80, 80);
            customBtn.Click += CustomColor_Click;
            Controls.Add(customBtn);

            // Opacity label
            var opLabel = new Label
            {
                Text = "Opacity",
                Location = new Point(12, 110),
                AutoSize = true,
                ForeColor = Color.FromArgb(200, 200, 200)
            };
            Controls.Add(opLabel);

            // Opacity track bar
            opacityTrack = new TrackBar
            {
                Minimum = 0,
                Maximum = 100,
                Value = Math.Clamp(currentOpacity, 0, 100),
                Location = new Point(12, 132),
                Size = new Size(200, 30),
                TickFrequency = 10,
                BackColor = Color.FromArgb(32, 32, 32)
            };
            opacityTrack.Scroll += (s, e) => UpdateOpacityLabel();
            Controls.Add(opacityTrack);

            // Opacity percentage label
            opacityLabel = new Label
            {
                Location = new Point(218, 135),
                AutoSize = true,
                ForeColor = Color.FromArgb(200, 200, 200)
            };
            Controls.Add(opacityLabel);
            UpdateOpacityLabel();

            // OK button
            var okBtn = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.OK,
                Location = new Point(120, 175),
                Size = new Size(80, 28),
                FlatStyle = FlatStyle.Flat,
                ForeColor = Color.White,
                BackColor = Color.FromArgb(55, 55, 55)
            };
            okBtn.FlatAppearance.BorderColor = Color.FromArgb(80, 80, 80);
            Controls.Add(okBtn);

            // Cancel button
            var cancelBtn = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Location = new Point(210, 175),
                Size = new Size(80, 28),
                FlatStyle = FlatStyle.Flat,
                ForeColor = Color.White,
                BackColor = Color.FromArgb(55, 55, 55)
            };
            cancelBtn.FlatAppearance.BorderColor = Color.FromArgb(80, 80, 80);
            Controls.Add(cancelBtn);

            AcceptButton = okBtn;
            CancelButton = cancelBtn;
        }

        private void SelectColor(Color color)
        {
            selectedColor = color;
            foreach (var p in swatchPanels)
                p.Invalidate();
        }

        private void SwatchPaint(object sender, PaintEventArgs e)
        {
            var panel = (Panel)sender;
            if (ColorsMatch(panel.BackColor, selectedColor))
            {
                using var pen = new Pen(Color.White, 2);
                e.Graphics.DrawRectangle(pen, 1, 1, panel.Width - 3, panel.Height - 3);
            }
        }

        private static bool ColorsMatch(Color a, Color b)
        {
            return a.R == b.R && a.G == b.G && a.B == b.B;
        }

        private void CustomColor_Click(object sender, EventArgs e)
        {
            using var dlg = new ColorDialog
            {
                Color = selectedColor,
                FullOpen = true
            };
            if (dlg.ShowDialog(this) == DialogResult.OK)
                SelectColor(dlg.Color);
        }

        private void UpdateOpacityLabel()
        {
            opacityLabel.Text = opacityTrack.Value + "%";
        }
    }
}
