using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
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
        private List<TreeNode> _favList;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void OnWindowLoaded(object sender, RoutedEventArgs e)
        {
            LoadExcelTree();
            LoadFavList();
        }

        private void OnWindowClosed(object sender, EventArgs e)
        {
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
                Content = treeNode.Name,
                Tag = treeNode.Path
            };
            bt.Click += FavItemClick;

            bt.ContextMenu = new ContextMenu();
            List<MenuItem> menuItems = new List<MenuItem>();
            bt.ContextMenu.ItemsSource = menuItems;
            var item = new MenuItem();
            item.Header = "删除";
            item.Click += MenuItemDeleteNodeClick;
            item.Tag = treeNode.Path;
            menuItems.Add(item);
            return bt;
        }

        private void LoadExcelTree()
        {
            string runningPath = AppDomain.CurrentDomain.BaseDirectory;
            string xlsPath = runningPath + "xls\\";
            xlsPath = "C:\\Work\\data\\xls\\";

            TreeNode rootNode = new TreeNode();
            Search(rootNode, xlsPath, NodeType.Dir);

            DataSource = rootNode.Child;
            DirTreeView.ItemsSource = DataSource;
        }

        private void Convert(object sender, RoutedEventArgs e)
        {

        }

        private void ScanDir(object sender, RoutedEventArgs e)
        {
            LoadExcelTree();
        }

        private void Search(TreeNode root, string path, NodeType nodeType)
        {
            root.Path = path;
            var lastIndexOf = path.LastIndexOf("\\", StringComparison.Ordinal) + 1;
            root.Name = path.Substring(lastIndexOf, path.Length - lastIndexOf);

            if (nodeType == NodeType.File)
            {
                root.Type = NodeType.File;
                return;
            }

            root.IsExpanded = false;
            root.Type = NodeType.Dir;
            var files = Directory.GetFiles(path, "*.xlsx", SearchOption.TopDirectoryOnly);
            root.ChildFileName = new List<string>(files);
            var childDirPath = Directory.GetDirectories(path, "*", SearchOption.TopDirectoryOnly);
            List<TreeNode> childNodes = new List<TreeNode>();
            for (int i = 0; i < childDirPath.Length; i++)
            {
                TreeNode node = new TreeNode();
                Search(node, childDirPath[i], NodeType.Dir);
                childNodes.Add(node);
            }

            for (int i = 0; i < files.Length; i++)
            {
                TreeNode node = new TreeNode();
                Search(node, files[i], NodeType.File);
                childNodes.Add(node);
            }

            root.Child = childNodes;
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

        private void OnItemSelect(object sender, RoutedEventArgs e)
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

        //public override int GetHashCode()
        //{
        //    int v1 = (int) Type;
        //    var charArray = Path.ToCharArray();
        //    int v2 = 0;
        //    for (int i = 0; i < charArray.Length; i++)
        //    {
        //        v2 += charArray[i];
        //    }
        //    return  +  + Name;
        //}
    }
}
