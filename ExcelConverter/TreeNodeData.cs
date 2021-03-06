﻿using System;
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


        //private List<string> _binNameList;

        //public List<string> BinNameList
        //{
        //    get
        //    {
        //        if (_binNameList == null)
        //        {
        //            _binNameList = new List<string>();
        //            if (SubSheetName != null)
        //            {
        //                for (int i = 0; i < SubSheetName.Count; i++)
        //                {
        //                    string binName = string.Empty;
        //                    Utils.GetBatCmd(SubSheetName[i], ref binName);
        //                    var arr = binName.Split(" ");
        //                    if(arr.Length >= 3)
        //                        _binNameList.Add(arr[2]);
        //                }
        //            }
        //        }

        //        return _binNameList;
        //    }
        //}

        public bool _expanded;
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
                if(!IsFile)
                {
                    _rev++;
                    for (var i = 0; i < Child.Count; i++)
                    {
                        Child[i].IsOn = value;
                    }
                    _rev--;
                }

                if(_rev == 0)
                    EventDispatcher.SendEvent(TaskType.NodeCheckedChanged, this);
            }
        }

        public static int _rev; //避免操作文件夹checkbox的value时 频繁reload TreeView
    }
}
