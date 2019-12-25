using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ExcelConverter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public List<TreeNode> DataSource { get; set; }
        private List<string> _favList = new List<string>();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void OnWindowLoaded(object sender, RoutedEventArgs e)
        {
            ReloadList();
        }

        private void OnWindowClosed(object sender, EventArgs e)
        {

        }

        private void ReloadList()
        {
            string runningPath = AppDomain.CurrentDomain.BaseDirectory;
            string xlsPath = runningPath + "xls\\";
            xlsPath = "C:\\Work\\data\\xls\\";

            TreeNode rootNode = new TreeNode();
            Search(rootNode, xlsPath, NodeType.Dir);

            DataSource = rootNode.Child;
            DirTreeView.ItemsSource = DataSource;

            //foreach (var item in this.DirTreeView.Items)
            //{
            //    DependencyObject dependencyObject = this.DirTreeView.ItemContainerGenerator.ContainerFromItem(item);
            //    if (dependencyObject != null)
            //    {
            //        ((TreeViewItem)dependencyObject).ExpandSubtree();
            //    }
            //}
        }

        private void Convert(object sender, RoutedEventArgs e)
        {

        }

        private void ScanDir(object sender, RoutedEventArgs e)
        {
            ReloadList();
        }

        private void Search(TreeNode root, string path, NodeType nodeType)
        {
            root.Path = path;
            root.Tag = root.GetHashCode();
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

        private void MenuItem_AddNode_Click(object sender, RoutedEventArgs e)
        {
            var tag = ((MenuItem)sender).Tag;
            FindNodeByTag(tag)
            var node = (TreeNode)treeItem.DataContext;
            _favList.Add(node.Path);

            Utils.SaveFav(_favList);
        }

        private void MenuItem_DeleteNode_Click(object sender, RoutedEventArgs e)
        {
            var treeItem = (TreeViewItem)sender;
            var node = (TreeNode)treeItem.DataContext;
            _favList.Remove(node.Path);

            Utils.SaveFav(_favList);
        }

        private void OnItemSelect(object sender, RoutedEventArgs e)
        {
            Button btn = (Button)sender;
            ItemContainerGenerator gen;
            TreeNode node = FindTreeNode((int)btn.Tag, out gen);

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
                Process.Start(new ProcessStartInfo(node.Path) { UseShellExecute = true }); ;
            }
        }

        private TreeNode FindTreeNode(int tag)
        {
            ItemContainerGenerator gen;
            return FindNodeByTag(DirTreeView.ItemContainerGenerator, tag, out gen);
        }
        
        private TreeNode FindTreeNode(int tag, out ItemContainerGenerator gen)
        {
            return FindNodeByTag(DirTreeView.ItemContainerGenerator, tag, out gen);
        }

        private TreeNode FindNodeByTag(ItemContainerGenerator container, int tag, out ItemContainerGenerator generator)
        {
            generator = null;
            foreach (TreeNode node in container.Items)
            {
                if (node.Tag == tag)
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

                    var findNode = FindNodeByTag(treeViewItem.ItemContainerGenerator, tag, out generator);
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
        public List<string> ChildFileName;
        public List<TreeNode> Child { get; set; }
        public bool IsExpanded { get; set; }
        public int Tag { get; set; }
        public NodeType Type;

        public string Path;

        public bool IsFile => Type == NodeType.File;

        public override string ToString()
        {
            return Name;
        }
    }
}
