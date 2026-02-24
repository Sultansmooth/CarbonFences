using System.Collections.Generic;

namespace CarbonZones.Model
{
    public class FenceTab
    {
        public string Name { get; set; } = "Tab";
        public List<string> Files { get; set; } = new List<string>();

        /// <summary>Tab accent color as ARGB int. 0 = inherit from fence.</summary>
        public int AccentColor { get; set; } = 0;
        /// <summary>Tab body background color as ARGB int. 0 = inherit from fence.</summary>
        public int BoxColor { get; set; } = 0;
        /// <summary>Tab label/header color as ARGB int. 0 = inherit from fence.</summary>
        public int LabelColor { get; set; } = 0;

        public FenceTab() { }
        public FenceTab(string name) { Name = name; }
        public FenceTab(string name, List<string> files) { Name = name; Files = files; }
    }
}
