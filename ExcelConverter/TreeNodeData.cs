using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace ExcelConverter
{
    public enum NodeType
    {
        Dir,
        File,
    }

    public partial class TreeNode
    {
        public string Name { get; set; }
        public List<string> ChildFileName { get; set; }
        public List<TreeNode> Child { get; set; }
        public bool IsExpanded { get; set; }
        public NodeType Type { get; set; }
        public string Path { get; set; }
        public bool IsMatch { get; set; }

        [JsonIgnore]
        public bool IsFile => Type == NodeType.File;
        [JsonIgnore]
        public System.Windows.Media.Brush Color
        {
            get
            {
                if (IsMatch)
                {
                    return System.Windows.Media.Brushes.LightGreen;
                }
                return System.Windows.Media.Brushes.White;
            }
        }
        [JsonIgnore]
        public string SingleFileName { get { return Utils.GetFileName(Path); } }
    }
}
