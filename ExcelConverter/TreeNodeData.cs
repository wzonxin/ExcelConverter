using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Text.Json.Serialization;
using System.Windows;

namespace ExcelConverter
{
    public enum NodeType
    {
        Dir,
        File,
    }

    public partial class TreeNode : INotifyPropertyChanged
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

        public bool _expanded;

        [JsonIgnore]
        public bool IsExpanded
        {
            get { return _expanded; }
            set { _expanded = value; }
        }

        [JsonIgnore] public bool IsMatch { get; set; }
        [JsonIgnore] public bool IsFile => Type == NodeType.File;

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
        public string SingleFileName
        {
            get { return Utils.GetFileName(Path); }
        }

        private bool _recordIsOn;

        [JsonIgnore]
        public bool IsOn
        {
            get { return _recordIsOn; }
            set
            {
                _recordIsOn = value;
                if (this.PropertyChanged != null)
                {
                    //通知ui 属性变化
                    this.PropertyChanged.Invoke(this, new PropertyChangedEventArgs("IsOn"));
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
    }
}