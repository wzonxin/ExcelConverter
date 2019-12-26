using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;

namespace ExcelConverter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public List<TreeNode> DataSource { get; set; }
        private List<TreeNode> _convertList = new List<TreeNode>();
        private List<TreeNode> _favList;
        private TreeNode _rootNode;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void OnWindowLoaded(object sender, RoutedEventArgs e)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);//注册Nuget包System.Text.Encoding.CodePages中的编码到.NET Core
            LoadExcelTree();
            LoadFavList();
        }

        private void OnWindowClosed(object sender, EventArgs e)
        {
        }

        private void RefreshConvertList()
        {
            if (_convertList.Count <= 0) return;

            float width = 70;
            float height = 30;
            int colMax = 6;

            ConvertGrid.Children.Clear();
            for (int i = 0; i < _convertList.Count; i++)
            {
                int row = i / colMax;
                int col = i % colMax;
                Button bt = GenConvertItem(height, _convertList[i]);

                Canvas.SetTop(bt, row * height + 5);
                Canvas.SetLeft(bt, col * width + 5);

                ConvertGrid.Children.Add(bt);
            }
        }

        private void LoadFavList()
        {
            _favList = Utils.ReadFav();
            if (_favList.Count <= 0) return;

            float width = 70;
            float height = 30;
            int colMax = 6;

            FavGrid.Children.Clear();
            for (int i = 0; i < _favList.Count; i++)
            {
                int row = i / colMax;
                int col = i % colMax;
                Button bt = GenFavItem(height, _favList[i]);

                Canvas.SetTop(bt, row * height + 5);
                Canvas.SetLeft(bt, col * width + 5);

                FavGrid.Children.Add(bt);
            }
        }

        private Button GenFavItem(float height, TreeNode treeNode)
        {
            Button bt = new Button()
            {
                Width = double.NaN,
                Height = height,
                Content = treeNode.SingleFileName,
                Tag = treeNode.Path,
                ToolTip = treeNode.BtnToolTip,
            };
            bt.Click += FavItemClick;

            bt.ContextMenu = new ContextMenu();
            List<MenuItem> menuItems = new List<MenuItem>();
            bt.ContextMenu.ItemsSource = menuItems;

            var item = new MenuItem();
            item.Header = "加入转表";
            item.Click += AddFavItemToCovertClick;
            item.Tag = treeNode.Path;
            menuItems.Add(item);

            item = new MenuItem();
            item.Header = "打开文件夹";
            item.Click += OpenFavItemFolder;
            item.Tag = treeNode.Path;
            menuItems.Add(item);

            item = new MenuItem();
            item.IsEnabled = false;
            item.Header = "————";
            menuItems.Add(item);

            item = new MenuItem();
            item.Header = "删除收藏";
            item.Click += MenuItemDeleteNodeClick;
            item.Tag = treeNode.Path;
            menuItems.Add(item);

            return bt;
        }

        private Button GenConvertItem(float height, TreeNode treeNode)
        {
            Button bt = new Button()
            {
                Width = double.NaN,
                Height = height,
                Content = treeNode.SingleFileName,
                ToolTip = treeNode.BtnToolTip,
            };

            bt.ContextMenu = new ContextMenu();
            List<MenuItem> menuItems = new List<MenuItem>();
            bt.ContextMenu.ItemsSource = menuItems;

            var item = new MenuItem();
            item.Header = "从列表中删除";
            item.Click += RemoveCovertItemClick;
            item.Tag = treeNode.Path;
            menuItems.Add(item);

            bt.Background = System.Windows.Media.Brushes.White;

            return bt;
        }

        private void LoadExcelTree()
        {
            TreeNode rootNode = Utils.ReadTree();
            _rootNode = rootNode;
            SetTreeSorce(rootNode);
        }

        private void Convert(object sender, RoutedEventArgs e)
        {
            Utils.ConvertExcel(_convertList);
        }

        private void SearchTextChange(object sender, TextChangedEventArgs e)
        {
            string inputText = SearchBox.Text;
            if (!string.IsNullOrEmpty(inputText))
            {
                TreeNode newTreeNode = null;
                Utils.FilterTree(_rootNode, inputText, ref newTreeNode);
                SetTreeSorce(newTreeNode);
                ExpandTree();
            }
            else
            {
                SetTreeSorce(_rootNode);
            }
        }

        private void ExpandTree()
        {
            foreach (var item in this.DirTreeView.Items)
            {
                DependencyObject dependencyObject = this.DirTreeView.ItemContainerGenerator.ContainerFromItem(item);
                if (dependencyObject != null)
                {
                    ((TreeViewItem)dependencyObject).ExpandSubtree();
                }
            }
        }

        private void SetTreeSorce(TreeNode node)
        {
            DataSource = node.Child;
            DirTreeView.ItemsSource = DataSource;
        }

        private void ScanDir(object sender, RoutedEventArgs e)
        {
            TreeNode rootNode = Utils.GenFileTree();
            _rootNode = rootNode;
            SetTreeSorce(rootNode);
        }

        private void MenuItemAddNodeClick(object sender, RoutedEventArgs e)
        {
            var tag = ((MenuItem)sender).Tag;
            var node = FindTreeNode((string) tag);
            if (!_favList.Contains(node))
            {
                _favList.Add(node);
                Utils.SaveFav(_favList);
            }
            LoadFavList();
        }

        private void AddFavItemToCovertClick(object sender, RoutedEventArgs e)
        {
            var tag = ((MenuItem)sender).Tag;
            var node = _favList.Find(treeNode => (string) tag == treeNode.Path);
            AddConvertNode(node);
        }

        private void AddConvertNode(TreeNode node)
        {
            if (node != null && !_convertList.Contains(node))
            {
                _convertList.Add(node);
                RefreshConvertList();
            }
        }

        private void RemoveCovertItemClick(object sender, RoutedEventArgs e)
        {
            var tag = ((MenuItem)sender).Tag;
            var node = _convertList.Find(treeNode => (string)tag == treeNode.Path);
            _convertList.Remove(node);

            RefreshConvertList();
        }

        private void MenuItemDeleteNodeClick(object sender, RoutedEventArgs e)
        {
            var tag = ((MenuItem)sender).Tag;
            var node = FindTreeNode((string) tag);
            if (_favList.Contains(node))
            {
                _favList.Remove(node);
                Utils.SaveFav(_favList);
            }
            LoadFavList();
        }

        private void TreeItemAddConvert(object sender, RoutedEventArgs e)
        {
            var tag = ((MenuItem)sender).Tag;
            var node = FindTreeNode((string)tag);

            AddConvertNode(node);
        }

        private void OpenFavItemFolder(object sender, RoutedEventArgs e)
        {
            var tag = ((MenuItem)sender).Tag;
            var node = _favList.Find(treeNode => (string)tag == treeNode.Path);
            OpenFolder(node);
        }
        
        private void OpenTreeItemFolder(object sender, RoutedEventArgs e)
        {
            var tag = ((MenuItem)sender).Tag;
            var node = FindTreeNode((string)tag);
            OpenFolder(node);
        }

        private static void OpenFolder(TreeNode node)
        {
            if (node != null)
            {
                var folderPath = node.Path.Remove(node.Path.LastIndexOf("\\"));
                Process.Start(new ProcessStartInfo(folderPath) { UseShellExecute = true });
            }
        }

        private void OnTreeItemSelect(object sender, RoutedEventArgs e)
        {
            Button btn = (Button)sender;
            ItemContainerGenerator gen;
            TreeNode node = FindTreeNode((string) btn.Tag, out gen);

            if (node == null)
                return;

            if (node.Type == NodeType.Dir)
            {
                DependencyObject dependencyObject = gen.ContainerFromItem(node);
                if (dependencyObject != null)
                {
                    var treeItem = ((TreeViewItem)dependencyObject);
                    treeItem.IsExpanded = !treeItem.IsExpanded;
                }
            }
            else
            {
                Process.Start(new ProcessStartInfo(node.Path) { UseShellExecute = true });
            }
        }

        private void FavItemClick(object sender, RoutedEventArgs e)
        {
            Button btn = (Button)sender;
            TreeNode node = FindTreeNode((string)btn.Tag);
            if (node == null)
                return;

            if (node.Type == NodeType.Dir)
            {
                Process.Start(new ProcessStartInfo(node.Path) { UseShellExecute = true });
            }
            else
            {
                Process.Start(new ProcessStartInfo(node.Path) { UseShellExecute = true });
            }
        }

        private TreeNode FindTreeNode(string tag)
        {
            ItemContainerGenerator gen;
            return FindNodeByTag(DirTreeView.ItemContainerGenerator, tag, out gen);
        }
        
        private TreeNode FindTreeNode(string path, out ItemContainerGenerator gen)
        {
            return FindNodeByTag(DirTreeView.ItemContainerGenerator, path, out gen);
        }

        private TreeNode FindNodeByTag(ItemContainerGenerator container, string path, out ItemContainerGenerator generator)
        {
            generator = null;
            foreach (TreeNode node in container.Items)
            {
                if (node.Path == path)
                {
                    generator = container;
                    return node;
                }
                else
                {
                    if (node.IsFile) continue;

                    TreeViewItem treeViewItem = (TreeViewItem)container.ContainerFromItem(node);
                    //第一次打开，列表没展开。会报空
                    if (treeViewItem == null)
                        continue;

                    var findNode = FindNodeByTag(treeViewItem.ItemContainerGenerator, path, out generator);
                    if (findNode != null)
                    {
                        return findNode;
                    }
                }
            }

            return null;
        }
    }

    public enum NodeType
    {
        Dir,
        File,
    }

    public class TreeNode
    {
        public string Name { get; set; }
        public List<string> ChildFileName { get; set; }
        public List<TreeNode> Child { get; set; }
        public bool IsExpanded { get; set; }
        //public int Tag { get; set; }
        public NodeType Type { get; set; }
        public string Path { get; set; }
        public bool IsFile => Type == NodeType.File;

        public string SingleFileName { get { return Utils.GetFileName(Path); } }
        public string BtnToolTip { 
            get
            {
                if(Name.Length > 60)
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
            } }

        public override string ToString()
        {
            return Name;
        }

        public override bool Equals(object? obj)
        {
            if (!(obj is TreeNode))
                return false;
            TreeNode other = (TreeNode) obj;
            return Type == other.Type && Path == other.Path && Name == other.Name;
        }
    }
}
