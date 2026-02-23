using System.Collections.Generic;

namespace CarbonZones.Model
{
    public class FenceTab
    {
        public string Name { get; set; } = "Tab";
        public List<string> Files { get; set; } = new List<string>();

        public FenceTab() { }
        public FenceTab(string name) { Name = name; }
        public FenceTab(string name, List<string> files) { Name = name; Files = files; }
    }
}
