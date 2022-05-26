using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;
using System.Windows;

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
        public List<string> SubSheetName { get; set; }
        public List<TreeNode> Child { get; set; }
        public NodeType Type { get; set; }
        public string Path { get; set; }
        public DateTime LastTime { get; set; }

        [JsonIgnore]
        public string WithSheetName
        {
            get { return GetWithSheetName(); } 
        }
        
        public bool _expanded = true;
        [JsonIgnore]
        public bool IsExpanded
        {
            get { return _expanded; } 
            set { _expanded = value; }
        }

        [JsonIgnore]
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

        private bool _recordIsOn;
        [JsonIgnore]
        public bool IsOn 
        {
            get
            {
                if (IsFile)
                    return _recordIsOn;
                return Child.TrueForAll(n => n.IsOn);
            }
            set
            {
                _recordIsOn = value;
                if (!IsFile)
                {
                    _rev++;
                    for (var i = 0; i < Child.Count; i++)
                    {
                        Child[i].IsOn = value;
                    }
                    _rev--;
                }

                if (_rev == 0)
                    EventDispatcher.SendEvent(TaskType.NodeCheckedChanged, this);
            }

            //get
            //{
            //    return _recordIsOn;
            //}
            //set
            //{
            //    _recordIsOn = value;
            //}
        }

        public static int _rev; //避免操作文件夹checkbox的value时 频繁reload TreeView
    }
}
