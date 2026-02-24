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

        private Color selectedAccent;
        private Color selectedLabel;
        private Color selectedBox;
        private readonly Panel[] swatchPanels;
        private readonly Panel labelPreview;
        private readonly Panel boxPreview;
        private readonly TrackBar opacityTrack;
        private readonly Label opacityLabel;

        public Color AccentColor => selectedAccent;
        public Color LabelColor => selectedLabel;
        public Color BoxColor => selectedBox;
        public int OpacityValue => opacityTrack.Value;

        public AppearanceDialog(Color currentAccent, int currentOpacity, Color currentLabel, Color currentBox)
        {
            selectedAccent = currentAccent;
            selectedLabel = currentLabel;
            selectedBox = currentBox;

            Text = "Appearance";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            Size = new Size(310, 370);
            BackColor = Color.FromArgb(32, 32, 32);
            ForeColor = Color.White;

            int y = 12;

            // ── Accent Color ──
            Controls.Add(new Label
            {
                Text = "Accent Color (tabs / border)",
                Location = new Point(12, y),
                AutoSize = true,
                ForeColor = Color.FromArgb(200, 200, 200)
            });
            y += 22;

            swatchPanels = new Panel[Presets.Length];
            for (int i = 0; i < Presets.Length; i++)
            {
                var swatch = new Panel
                {
                    Size = new Size(28, 28),
                    Location = new Point(12 + (i % 5) * 34, y + (i / 5) * 34),
                    BackColor = Presets[i],
                    Cursor = Cursors.Hand
                };
                int idx = i;
                swatch.Click += (s, e) => SelectAccent(Presets[idx]);
                swatch.Paint += SwatchPaint;
                swatchPanels[i] = swatch;
                Controls.Add(swatch);
            }

            var customBtn = MakeButton("Custom...", new Point(190, y), new Size(90, 28));
            customBtn.Click += (s, e) => PickColor(selectedAccent, c => SelectAccent(c));
            Controls.Add(customBtn);

            y += 74;

            // ── Label Color ──
            Controls.Add(new Label
            {
                Text = "Label Color (title bar)",
                Location = new Point(12, y),
                AutoSize = true,
                ForeColor = Color.FromArgb(200, 200, 200)
            });
            y += 22;

            labelPreview = new Panel
            {
                Size = new Size(28, 28),
                Location = new Point(12, y),
                BackColor = selectedLabel,
                BorderStyle = BorderStyle.FixedSingle
            };
            Controls.Add(labelPreview);

            var labelPickBtn = MakeButton("Pick...", new Point(48, y), new Size(70, 28));
            labelPickBtn.Click += (s, e) => PickColor(selectedLabel, c => { selectedLabel = c; labelPreview.BackColor = c; });
            Controls.Add(labelPickBtn);

            var labelResetBtn = MakeButton("Reset", new Point(124, y), new Size(60, 28));
            labelResetBtn.Click += (s, e) => { selectedLabel = Color.Black; labelPreview.BackColor = Color.Black; };
            Controls.Add(labelResetBtn);

            y += 38;

            // ── Box Color ──
            Controls.Add(new Label
            {
                Text = "Box Color (background)",
                Location = new Point(12, y),
                AutoSize = true,
                ForeColor = Color.FromArgb(200, 200, 200)
            });
            y += 22;

            boxPreview = new Panel
            {
                Size = new Size(28, 28),
                Location = new Point(12, y),
                BackColor = selectedBox,
                BorderStyle = BorderStyle.FixedSingle
            };
            Controls.Add(boxPreview);

            var boxPickBtn = MakeButton("Pick...", new Point(48, y), new Size(70, 28));
            boxPickBtn.Click += (s, e) => PickColor(selectedBox, c => { selectedBox = c; boxPreview.BackColor = c; });
            Controls.Add(boxPickBtn);

            var boxResetBtn = MakeButton("Reset", new Point(124, y), new Size(60, 28));
            boxResetBtn.Click += (s, e) => { selectedBox = Color.Black; boxPreview.BackColor = Color.Black; };
            Controls.Add(boxResetBtn);

            y += 38;

            // ── Opacity ──
            Controls.Add(new Label
            {
                Text = "Opacity",
                Location = new Point(12, y),
                AutoSize = true,
                ForeColor = Color.FromArgb(200, 200, 200)
            });
            y += 22;

            opacityTrack = new TrackBar
            {
                Minimum = 0,
                Maximum = 100,
                Value = Math.Clamp(currentOpacity, 0, 100),
                Location = new Point(12, y),
                Size = new Size(200, 30),
                TickFrequency = 10,
                BackColor = Color.FromArgb(32, 32, 32)
            };
            opacityTrack.Scroll += (s, e) => UpdateOpacityLabel();
            Controls.Add(opacityTrack);

            opacityLabel = new Label
            {
                Location = new Point(218, y + 3),
                AutoSize = true,
                ForeColor = Color.FromArgb(200, 200, 200)
            };
            Controls.Add(opacityLabel);
            UpdateOpacityLabel();

            y += 40;

            // ── OK / Cancel ──
            var okBtn = MakeButton("OK", new Point(120, y), new Size(80, 28));
            okBtn.DialogResult = DialogResult.OK;
            Controls.Add(okBtn);

            var cancelBtn = MakeButton("Cancel", new Point(210, y), new Size(80, 28));
            cancelBtn.DialogResult = DialogResult.Cancel;
            Controls.Add(cancelBtn);

            AcceptButton = okBtn;
            CancelButton = cancelBtn;
        }

        private void SelectAccent(Color color)
        {
            selectedAccent = color;
            foreach (var p in swatchPanels)
                p.Invalidate();
        }

        private void SwatchPaint(object sender, PaintEventArgs e)
        {
            var panel = (Panel)sender;
            if (ColorsMatch(panel.BackColor, selectedAccent))
            {
                using var pen = new Pen(Color.White, 2);
                e.Graphics.DrawRectangle(pen, 1, 1, panel.Width - 3, panel.Height - 3);
            }
        }

        private static bool ColorsMatch(Color a, Color b)
        {
            return a.R == b.R && a.G == b.G && a.B == b.B;
        }

        private void PickColor(Color current, Action<Color> onPicked)
        {
            using var dlg = new ColorDialog
            {
                Color = current,
                FullOpen = true
            };
            if (dlg.ShowDialog(this) == DialogResult.OK)
                onPicked(dlg.Color);
        }

        private void UpdateOpacityLabel()
        {
            opacityLabel.Text = opacityTrack.Value + "%";
        }

        private static Button MakeButton(string text, Point location, Size size)
        {
            var btn = new Button
            {
                Text = text,
                Location = location,
                Size = size,
                FlatStyle = FlatStyle.Flat,
                ForeColor = Color.White,
                BackColor = Color.FromArgb(55, 55, 55)
            };
            btn.FlatAppearance.BorderColor = Color.FromArgb(80, 80, 80);
            return btn;
        }
    }
}
