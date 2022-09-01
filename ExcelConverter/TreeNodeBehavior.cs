using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace ExcelConverter
{
    public partial class TreeNode
    {
        public TreeNode()
        {
            Path = string.Empty;
            Name = string.Empty;
        }

        public string GetBtnToolTip()
        {
            var withSheetName = GetWithSheetName();
            if (withSheetName.Length > 60)
            {
                var spilt = withSheetName.Replace(SingleFileName, "").Split(",");
                string s = "";
                for (int i = 0; i < spilt.Length; i++)
                {
                    s += spilt[i] + "\n";
                }
                return s;
            }
            return withSheetName;
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

        public TreeNode Clone(bool includeChild = true)
        {
            TreeNode cloneNode = new TreeNode();
            cloneNode.Name = Name;
            cloneNode.Path = Path;
            cloneNode.IsExpanded = IsExpanded;
            cloneNode.Type = Type;
            if (SubSheetName != null)
            {
                cloneNode.SubSheetName = new List<string>(SubSheetName);
            }

            if (includeChild)
            {
                if (Child != null)
                {
                    cloneNode.Child = new List<TreeNode>(Child.Count);
                    for (var i = 0; i < Child.Count; i++)
                    {
                        var childCloneNode = Child[i].Clone();
                        cloneNode.Child.Add(childCloneNode);
                    }
                }
            }
            cloneNode.IsOn = IsOn;
            return cloneNode;
        }

        public void AutoOpen()
        {
            string absolutePath = GetAbsolutePath();
            if (IsFile)
            {
                if (System.IO.File.Exists(absolutePath))
                    OpenFolderFile();
            }
            else
            {
                if(System.IO.Directory.Exists(absolutePath))
                    OpenFolderDir();
            }
        }

        public void OpenFolderDir()
        {
            Process.Start(new ProcessStartInfo(GetAbsolutePath()) { UseShellExecute = true });
        }

        public void OpenFolderFile()
        {
            Process.Start("Explorer.exe", "/select," + GetAbsolutePath());
        }

        public void OpenFile()
        {
            var path = GetAbsolutePath();
            if (System.IO.File.Exists(path))
                Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
        }

        public string GetAbsolutePath()
        {
            return Utils.GetAbsolutePath(Path);
        }

        private string GenXlsName()
        {
            if (SubSheetName == null)
                return string.Empty;

            string sheetNames = "(";
            int sheetCnt = SubSheetName.Count;
            for (int j = 0; j < sheetCnt; j++)
            {
                var sheetName = SubSheetName[j];
                sheetNames += sheetName;
                if (j < sheetCnt - 1)
                    sheetNames += ", ";
            }
            sheetNames += ")";
            return sheetNames;
        }

        private string _name;
        public string GetWithSheetName()
        {
            if (string.IsNullOrEmpty(_name))
            {
                if (IsFile)
                    _name = Name + GenXlsName();
                else
                    _name = Name;
            }

            return _name;
        }

        public bool MatchSearch(string searchContent)
        {
            string childName = GetWithSheetName();
            if (childName.Contains(searchContent, StringComparison.CurrentCultureIgnoreCase))
            {
                return true;
            }

            //if (BinNameList != null)
            //{
            //    for (int i = 0; i < BinNameList.Count; i++)
            //    {
            //        if (BinNameList[i].Contains(searchContent))
            //            return true;
            //    }
            //}

            return false;
        }

        public void Recursive(Action<TreeNode> call)
        {
            call(this);

            if (!IsFile)
            {
                if (Child == null)
                {
                    return;
                }

                for (int i = 0; i < Child.Count; i++)
                {
                    Child[i].Recursive(call);
                }
            }
        }

        public static bool _banChildChange;
        public void SetChecked(bool isOn)
        {
            IsOn = isOn;

            if (_banChildChange)
            {
                return;
            }

            if (!IsFile)
            {
                for (int i = 0; i < Child.Count; i++)
                {
                    Child[i].SetChecked(isOn);
                }
            }
        }

        public bool CheckChildAllOn()
        {
            bool allOn = true;
            for (int i = 0; i < Child.Count; i++)
            {
                if (!Child[i].IsOn)
                {
                    allOn = false;
                    break;
                }
            }
            return allOn;
        }

        public TreeNode FindNodeInChild(string tag)
        {
            if (Path == tag)
            {
                return this;
            }

            if (!IsFile && Child != null)
            {
                for (int i = 0; i < Child.Count; i++)
                {
                    var findNodeInChild = Child[i].FindNodeInChild(tag);
                    if (findNodeInChild != null)
                    {
                        return findNodeInChild;
                    }
                }
            }

            return null;

        }
    }
}
