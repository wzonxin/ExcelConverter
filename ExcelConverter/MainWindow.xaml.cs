using System;
using System.Collections.Generic;
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
        private System.Windows.Threading.DispatcherTimer _timer;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void OnWindowLoaded(object sender, RoutedEventArgs e)
        {
            Utils.InitWorkingPath();
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);//注册Nuget包System.Text.Encoding.CodePages中的编码到.NET Core
            LoadExcelTree();
            LoadFavList();

            _timer = new System.Windows.Threading.DispatcherTimer();
            _timer.Interval = TimeSpan.FromMilliseconds(100);
            // Set the callback to just show the time ticking away
            // NOTE: We are using a control so this has to run on 
            // the UI thread
            _timer.Tick += new EventHandler(TimerTick);
            _timer.Start();

            InitEvent();
        }

    private void InitEvent()
    {
        EventDispatcher.RegdEvent<string>(TaskType.SearchError, OnSearchError);
        EventDispatcher.RegdEvent<float>(TaskType.UpdateSearchProgress, UpdateProgress);
        EventDispatcher.RegdEvent<TreeNode>(TaskType.FinishedSearch, OnFinishedSearch);
        EventDispatcher.RegdEvent(TaskType.NodeCheckedChanged, NodeCheckedChanged);
    }

    private void TimerTick(object s, EventArgs a)
        {
            EventDispatcher.CheckTick();
        }

        private void OnWindowClosed(object sender, EventArgs e)
        {
            _timer?.Stop();
            EventDispatcher.Clear();
        }

        private void RefreshConvertList()
        {
            float width = 70;
            float height = 30;
            int colMax = 6;

            ConvertGrid.Children.Clear();
            for (int i = 0; i < _convertList.Count; i++)
            {
                int row = i / colMax;
                int col = i % colMax;
                Button bt = GenConvertItem(height, _convertList[i]);

                Button removeBt = new Button();
                removeBt.Tag = bt.Tag;
                removeBt.Click += RemoveCovertBtnClick;
                removeBt.Content = "X";
                removeBt.Width = 20;
                removeBt.Height = height - 10;

                Canvas.SetTop(bt, row * height + 5);
                Canvas.SetLeft(bt, col * width + 5);

                Canvas.SetTop(removeBt, row * height + 5);
                Canvas.SetLeft(removeBt, row * width + 5 + 50);

                Panel panel = new WrapPanel();
                panel.Children.Add(bt);
                panel.Children.Add(removeBt);

                ConvertGrid.Children.Add(panel);
            }
        }

        private void LoadFavList()
        {
            _favList = Utils.ReadFav();

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
                ToolTip = treeNode.GetBtnToolTip(),
                Background = treeNode.GetBtnColor(),
            };
            bt.Click += FavBtnItemClick;

            bt.ContextMenu = new ContextMenu();
            List<MenuItem> menuItems = new List<MenuItem>();
            bt.ContextMenu.ItemsSource = menuItems;

            var item = new MenuItem();
            item.Header = "直接转表";
            item.Click += (obj,ev) =>
            {
                Utils.ConvertExcel(new List<TreeNode>(){ treeNode });
            };
            item.Tag = treeNode.Path;
            menuItems.Add(item);
            
            //item = new MenuItem();
            //item.Header = "加入待转";
            //item.Click += AddFavItemToCovertClick;
            //item.Tag = treeNode.Path;
            //menuItems.Add(item);

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
                Width = 90,
                Height = height,
                Content = treeNode.SingleFileName,
                ToolTip = treeNode.GetBtnToolTip(),
                Background = treeNode.GetBtnColor(),
                Tag = treeNode.Path,
            };
            bt.Click += ConvertNodeBtnItemClick;

            bt.ContextMenu = new ContextMenu();
            List<MenuItem> menuItems = new List<MenuItem>();
            bt.ContextMenu.ItemsSource = menuItems;

            var item = new MenuItem();
            item.Header = "从列表中删除";
            item.Click += RemoveCovertItemClick;
            item.Tag = treeNode.Path;
            menuItems.Add(item);
            
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
            if (_convertList.Count <= 0)
            {
                return;
            }

            Point point = new Point(Left, Top);
            Utils.SetConsolePos(point);
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
                ClearSearchBtn.Visibility = Visibility.Visible;
            }
            else
            {
                SetTreeSorce(_rootNode);
                ClearSearchBtn.Visibility = Visibility.Hidden;
            }
        }

        private void ClickClearSearchBtn(object sender, RoutedEventArgs e)
        {
            SearchBox.Text = string.Empty;
        }

        private void SetTreeSorce(TreeNode node)
        {
            DataSource = node.Child;
            DirTreeView.ItemsSource = DataSource;
        }

        private void ScanDir(object sender, RoutedEventArgs e)
        {
            //TreeNode rootNode = Utils.GenFileTree(ScanProgressBar, UpdateProgress);
            //_rootNode = rootNode;
            //SetTreeSorce(rootNode);
            ScanLabel.Content = "扫描中...";
            Utils.GenFileTree();
        }

        private void OnClearConvertList(object sender, RoutedEventArgs e)
        {
            TreeNode._rev = 1;
            _convertList.Clear();
            RefreshConvertList();
            ClearToggleNode(_rootNode);
            TreeNode._rev = 0;
            EventDispatcher.SendEvent(TaskType.NodeCheckedChanged);
        }

        private void ClearToggleNode(TreeNode checkNode)
        {
            if (checkNode.IsFile)
            {
                checkNode.IsOn = false;
            }
            else
            {
                for (int i = 0; i < checkNode.Child.Count; i++)
                {
                    TreeNode node = checkNode.Child[i];
                    ClearToggleNode(node);
                }
            }
        }

        private void UpdateProgress(float value)
        {
            float val = value;
            ScanProgressBar.Value = val * 100;
        }

        private void OnFinishedSearch(TreeNode node)
        {
            _rootNode = node;
            SetTreeSorce(node);
            ScanLabel.Content = "扫描完成";
        }

        private void OnSearchError(string errorStr)
        {
            ScanLabel.Content = errorStr;
        }

        private void NodeCheckedChanged()
        {
            DirTreeView.ItemsSource = null;
            SetTreeSorce(_rootNode);

            _convertList.Clear();
            CollectToggleNode(_rootNode);
            _convertList.Sort(Utils.SortList);
            RefreshConvertList();
        }

        private void CollectToggleNode(TreeNode checkNode)
        {
            if (checkNode.IsFile)
            {
                if (checkNode.IsOn)
                {
                    _convertList.Add(checkNode);
                }
            }
            else
            {
                for (int i = 0; i < checkNode.Child.Count; i++)
                {
                    TreeNode node = checkNode.Child[i];
                    CollectToggleNode(node);
                }
            }
        }

        private void MenuItemAddNodeClick(object sender, RoutedEventArgs e)
        {
            var tag = ((MenuItem)sender).Tag;
            var node = FindTreeNode((string) tag);
            if (!_favList.Contains(node))
            {
                _favList.Add(node);
                _favList.Sort(Utils.SortList);
                Utils.SaveFav(_favList);
            }
            LoadFavList();
        }

        private void AddFavItemToCovertClick(object sender, RoutedEventArgs e)
        {
            var tag = ((MenuItem)sender).Tag;
            var node = _favList.Find(treeNode => (string) tag == treeNode.Path);
            //AddConvertNode(node);
        }

        //private void AddConvertNode(TreeNode node)
        //{
        //    if (node != null && !_convertList.Contains(node))
        //    {
        //        _convertList.Add(node);
        //        _convertList.Sort(Utils.SortList);
        //        RefreshConvertList();
        //    }
        //}

        private void RemoveCovertItemClick(object sender, RoutedEventArgs e)
        {
            var menuItem = (MenuItem)sender;
            var tag = menuItem.Tag;
            RemoveConvertNode(tag);
        }
        
        private void RemoveCovertBtnClick(object sender, RoutedEventArgs e)
        {
            var button = (Button)sender;
            var tag = button.Tag;
            RemoveConvertNode(tag);
        }

        private void RemoveConvertNode(object tag)
        {
            var node = _convertList.Find(treeNode => (string) tag == treeNode.Path);
            _convertList.Remove(node);
            RefreshConvertList();
            node.IsOn = false;
            EventDispatcher.SendEvent(TaskType.NodeCheckedChanged);
        }

        private void MenuItemDeleteNodeClick(object sender, RoutedEventArgs e)
        {
            var tag = ((MenuItem)sender).Tag;
            var node = _favList.Find((node)=> (string) tag == node.Path);
            if (node != null)
            {
                _favList.Remove(node);
                Utils.SaveFav(_favList);
            }
            LoadFavList();
        }

        private void TreeItemDirectConvert(object sender, RoutedEventArgs e)
        {
            var tag = ((MenuItem)sender).Tag;
            var node = FindTreeNode((string)tag);

            Utils.ConvertExcel(new List<TreeNode>(){ node });
        }
        
        private void TreeItemAddConvert(object sender, RoutedEventArgs e)
        {
            var tag = ((MenuItem)sender).Tag;
            var node = FindTreeNode((string)tag);

            //AddConvertNode(node);
        }

        private void OpenFavItemFolder(object sender, RoutedEventArgs e)
        {
            var tag = ((MenuItem)sender).Tag;
            var node = _favList.Find(treeNode => (string)tag == treeNode.Path);
            node.AutoOpen();
        }

        private void OpenTreeItemFolder(object sender, RoutedEventArgs e)
        {
            var tag = ((MenuItem)sender).Tag;
            var node = FindTreeNode((string)tag);
            node.AutoOpen();
        }

        private void OnTreeItemBtnSelect(object sender, RoutedEventArgs e)
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
                    //node.IsExpanded = treeItem.IsExpanded;
                }
            }
            else
            {
                node.OpenFile();
            }
        }

        private void FavBtnItemClick(object sender, RoutedEventArgs e)
        {
            Button btn = (Button)sender;
            //TreeNode node = FindTreeNode((string)btn.Tag);
            //if (node == null)
            //    return;

            var tag = (string) btn.Tag;
            var node = _favList.Find(treeNode => (string)tag == treeNode.Path);

            if (node.IsFile)
            {
                node.OpenFile();
            }
            else
            {
                node.OpenFolderDir();
            }
        }

        
        private void ConvertNodeBtnItemClick(object sender, RoutedEventArgs e)
        {
            Button btn = (Button)sender;
            //TreeNode node = FindTreeNode((string)btn.Tag);
            //if (node == null)
            //    return;

            var tag = (string) btn.Tag;
            var node = _convertList.Find(treeNode => (string)tag == treeNode.Path);

            if (node.IsFile)
            {
                node.OpenFile();
            }
            else
            {
                node.OpenFolderDir();
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

        private void OnTreeViewItemCollapseStateChanged(object sender, RoutedEventArgs e)
        {
            var treeViewItem = (TreeViewItem)e.Source;
            var node = treeViewItem.DataContext as TreeNode;
            if (node != null)
            {
                node.IsExpanded = treeViewItem.IsExpanded;
            }
        }
    }
}
