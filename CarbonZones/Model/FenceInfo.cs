using System;
using System.Collections.Generic;

namespace CarbonZones.Model
{
    public class FenceInfo
    {
        /* 
         * DO NOT RENAME PROPERTIES. Used for XML serialization.
         */

        public Guid Id { get; set; }

        public string Name { get; set; }

        public int PosX { get; set; }

        public int PosY { get; set; }

        /// <summary>
        /// Gets or sets the DPI scaled window width.
        /// </summary>
        public int Width { get; set; }

        /// <summary>
        /// Gets or sets the DPI scaled window height.
        /// </summary>
        public int Height { get; set; }

        public bool Locked { get; set; }

        public bool CanMinify { get; set; }

        /// <summary>
        /// Gets or sets the logical window title height.
        /// </summary>
        public int TitleHeight { get; set; } = 35;

        /// <summary>
        /// Accent color as ARGB int. Default is blue (100, 160, 230).
        /// </summary>
        public int AccentColor { get; set; } = -10182682; // Color.FromArgb(100, 160, 230).ToArgb()

        /// <summary>
        /// Background opacity 0–100. Default 60.
        /// </summary>
        public int Opacity { get; set; } = 60;

        /// <summary>
        /// Title bar background color as ARGB int. 0 = default black.
        /// </summary>
        public int LabelColor { get; set; } = 0;

        /// <summary>
        /// Fence body background color as ARGB int. 0 = default black.
        /// </summary>
        public int BoxColor { get; set; } = 0;

        /// <summary>
        /// Icon size in pixels (width). Default 75. Range 50–150.
        /// </summary>
        public int IconSize { get; set; } = 75;

        public List<string> Files { get; set; } = new List<string>();

        public List<FenceTab> Tabs { get; set; } = new List<FenceTab>();

        public FenceInfo()
        {

        }

        public FenceInfo(Guid id)
        {
            Id = id;
        }
    }
}
