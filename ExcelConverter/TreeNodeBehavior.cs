using System.Collections.Generic;
using System.Diagnostics;

namespace ExcelConverter
{
    public partial class TreeNode
    {
        public string GetBtnToolTip()
        {
            if (Name.Length > 60)
            {
                var spilt = Name.Replace(SingleFileName, "").Split(",");
                string s = "";
                for (int i = 0; i < spilt.Length; i++)
                {
                    s += spilt[i] + "\n";
                }
                return s;
            }
            return Name;
        }

        public System.Windows.Media.SolidColorBrush GetBtnColor()
        {
            if (IsFile)
            {
               return System.Windows.Media.Brushes.GhostWhite;
            }
            else
            {
                return System.Windows.Media.Brushes.LightGoldenrodYellow;
            }
        }

        public override string ToString()
        {
            return Name;
        }

        public override bool Equals(object obj)
        {
            if (!(obj is TreeNode))
                return false;
            TreeNode other = (TreeNode)obj;
            return Type == other.Type && Path == other.Path && Name == other.Name;
        }

        public TreeNode Clone()
        {
            TreeNode treeNode = new TreeNode();
            treeNode.Name = Name;
            treeNode.Path = Path;
            treeNode.IsExpanded = IsExpanded;
            treeNode.Type = Type;
            if (ChildFileName != null)
            {
                treeNode.ChildFileName = new List<string>(ChildFileName);
            }
            if (Child != null)
            {
                treeNode.Child = new List<TreeNode>(Child);
            }
            return treeNode;
        }

        public void AutoOpen()
        {
            if (IsFile)
                OpenFolderFile();
            else
                OpenFolderDir();
        }

        public void OpenFolderDir()
        {
            Process.Start(new ProcessStartInfo(Path) { UseShellExecute = true });
        }

        public void OpenFolderFile()
        {
            Process.Start("Explorer.exe", "/select," + Path);
        }

        public void OpenFile()
        {
            Process.Start(new ProcessStartInfo(Path) { UseShellExecute = true });
        }
    }
}
