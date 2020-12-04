using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
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

        private TreeNode _curTreeNode
        {
            get
            {
                bool inSearch = !string.IsNullOrEmpty(SearchBox.Text);
                if (inSearch)
                    return _searchTreeNode;

                return _rootNode;
            }
        }

        public MainWindow()
        {
            InitializeComponent();
        }

        private void OnWindowLoaded(object sender, RoutedEventArgs e)
        {
            Utils.InitWorkingPath();
            ServerLogChecker.InitServerLogPath();
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);//注册Nuget包System.Text.Encoding.CodePages中的编码到.NET Core
            ServerLogChecker.InitBinExcelMap();
            LoadExcelTree();
            LoadFavList();
            InitComboBox();

            _timer = new System.Windows.Threading.DispatcherTimer();
            _timer.Interval = TimeSpan.FromMilliseconds(1);
            // Set the callback to just show the time ticking away
            // NOTE: We are using a control so this has to run on 
            // the UI thread
            _timer.Tick += new EventHandler(TimerTick);
            _timer.Start();

            InitEvent();
            LocationChanged += SyncPopupPosition;
        }

        private void SyncPopupPosition(object sender, EventArgs e)
        {
            var offset = SearchBinPopup.HorizontalOffset;
            SearchBinPopup.HorizontalOffset = offset + 1;
            SearchBinPopup.HorizontalOffset = offset;
        }

        private void InitEvent()
        {
            EventDispatcher.RegdEvent<string>(TaskType.SearchError, OnSearchError);
            EventDispatcher.RegdEvent<float>(TaskType.UpdateSearchProgress, UpdateProgress);
            EventDispatcher.RegdEvent<TreeNode>(TaskType.FinishedSearch, OnFinishedSearch);
            EventDispatcher.RegdEvent<TreeNode>(TaskType.NodeCheckedChanged, NodeCheckedChanged);
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

        private void OnWindowDeactivated(object sender, EventArgs e)
        {
            HidePopupClick(null, null);
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
                removeBt.Tag = _convertList[i].Path;
                removeBt.Click += RemoveCovertBtnClick;
                removeBt.Content = "X";
                removeBt.Width = 20;
                removeBt.Height = height - 10;

                //Canvas.SetTop(bt, row * height + 5);
                //Canvas.SetLeft(bt, col * width + 5);

                //Canvas.SetTop(removeBt, row * height + 5);
                //Canvas.SetLeft(removeBt, row * width + 5 + 50);

                StackPanel panel = new StackPanel();
                panel.Orientation = Orientation.Horizontal;
                panel.Width = 120;
                panel.Height = height + 6;
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
            item.Click += (obj, ev) =>
            {
                DoConvert(new List<TreeNode>() { treeNode });
            };
            item.Tag = treeNode.Path;
            menuItems.Add(item);

            item = new MenuItem();
            item.Header = "加入待转";
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
                Width = 90,
                Height = height,
                Content = treeNode.SingleFileName,
                ToolTip = treeNode.GetBtnToolTip(),
                Background = treeNode.GetBtnColor(),
                Tag = treeNode.Path,
            };
            bt.Click += ConvertNodeBtnItemClick;

            //bt.ContextMenu = new ContextMenu();
            //List<MenuItem> menuItems = new List<MenuItem>();
            //bt.ContextMenu.ItemsSource = menuItems;

            //var item = new MenuItem();
            //item.Header = "从列表中删除";
            //item.Click += RemoveCovertItemClick;
            //item.Tag = treeNode.Path;
            //menuItems.Add(item);

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

            DoConvert(_convertList);
        }

        private void OneKeyAdd(object sender, RoutedEventArgs e)
        {
            var list = Utils.GetModifyList();
            TreeNode._rev = 1;

            for (var i = 0; i < list.Count; i++)
            {
                var node = _rootNode.FindNodeInChild(list[i]);
                if (node != null && !DirFilter.IsSkipDir(node.Path))
                {
                    node.IsOn = true;
                }
            }

            TreeNode._rev = 0;
            NodeCheckedChanged(null);
        }

        private void DoConvert(List<TreeNode> nodeList)
        {
            ConvertDialog dlg = new ConvertDialog();
            dlg.Owner = this;
            dlg.OnClosedEvent += Utils.CleanConvert;
            dlg.Show();

            Utils.ConvertExcel(nodeList);

        }

        private void SearchTextChange(object sender, TextChangedEventArgs e)
        {
            string inputText = SearchBox.Text;
            SyncTreeChecked();
            if (!string.IsNullOrEmpty(inputText))
            {
                _searchTreeNode = null;
                Utils.FilterTree(_rootNode, inputText, ref _searchTreeNode);
                SetTreeSorce(_searchTreeNode);
            }
            else
            {
                SetTreeSorce(_rootNode);
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
            if (_searchTreeNode != null)
                ClearToggleNode(_searchTreeNode);
            TreeNode._rev = 0;
            NodeCheckedChanged(null);
        }

        private void ClearToggleNode(TreeNode checkNode)
        {
            checkNode.Recursive(node => { node.IsOn = false; });
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
            SearchTextChange(null, null);

            MessageBox.Show("扫描完成");
        }

        private void OnSearchError(string errorStr)
        {
            //ScanLabel.Content = errorStr;
            MessageBox.Show($"{errorStr}");
        }

        private void NodeCheckedChanged(TreeNode changedNode)
        {
            SyncTreeChecked(changedNode);

            TreeNode treeRootNode = _curTreeNode;
            bool inSearch = !string.IsNullOrEmpty(SearchBox.Text);

            DirTreeView.ItemsSource = null;
            SetTreeSorce(treeRootNode);

            _convertList.Clear();

            CollectToggleNode(_rootNode);
            if (inSearch)
            {
                CollectToggleNode(_searchTreeNode);
            }

            _convertList.Sort(Utils.SortList);
            RefreshConvertList();
        }

        //不传参数就是全量更新
        private void SyncTreeChecked(TreeNode updateNode = null)
        {
            List<TreeNode> updateList = new List<TreeNode>();
            if (updateNode != null)
                updateList.Add(updateNode);
            else
                updateNode = _searchTreeNode;

            if (updateNode == null)
                return;

            updateNode.Recursive(node =>
            {
                if (node.IsFile)
                {
                    if (!updateList.Contains(node))
                        updateList.Add(node);
                }
            });

            _rootNode.Recursive(node =>
            {
                for (int i = 0; i < updateList.Count; i++)
                {
                    var checkNode = updateList[i];
                    if (checkNode.Name == node.Name && checkNode.Path == node.Path)
                    {
                        node.JustSetChecked(checkNode.IsOn);
                    }
                }
            });
        }

        private void CollectToggleNode(TreeNode checkNode)
        {
            checkNode.Recursive(node =>
            {
                if (node.IsOn)
                {
                    if (!_convertList.Contains(node))
                        _convertList.Add(node);
                }
            });
        }

        private void MenuItemAddNodeClick(object sender, RoutedEventArgs e)
        {
            var tag = ((MenuItem)sender).Tag;
            var node = FindTreeViewNode((string)tag);
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
            var node = FindNodeInDataTree((string)tag);
            node.IsOn = true;
        }

        //private void RemoveCovertItemClick(object sender, RoutedEventArgs e)
        //{
        //    var menuItem = (MenuItem)sender;
        //    var tag = menuItem.Tag;
        //    RemoveConvertNode(tag);
        //}

        private void RemoveCovertBtnClick(object sender, RoutedEventArgs e)
        {
            var button = (Button)sender;
            var tag = button.Tag;
            RemoveConvertNode(tag);
        }

        private void RemoveConvertNode(object tag)
        {
            var node = _convertList.Find(treeNode => (string)tag == treeNode.Path);
            _convertList.Remove(node);
            RefreshConvertList();
            var viewNode = FindNodeInDataTree((string)tag);
            if (viewNode != null)
            {
                viewNode.IsOn = false;
                EventDispatcher.SendEvent(TaskType.NodeCheckedChanged, viewNode);
            }
        }

        private void MenuItemDeleteNodeClick(object sender, RoutedEventArgs e)
        {
            var tag = ((MenuItem)sender).Tag;
            var node = _favList.Find((node) => (string)tag == node.Path);
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
            var node = FindTreeViewNode((string)tag);

            DoConvert(new List<TreeNode>() { node });
        }

        private void TreeItemAddConvert(object sender, RoutedEventArgs e)
        {
            var tag = ((MenuItem)sender).Tag;
            var node = FindTreeViewNode((string)tag);

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
            var node = FindTreeViewNode((string)tag);
            node.AutoOpen();
        }

        private void OnTreeItemBtnSelect(object sender, RoutedEventArgs e)
        {
            Button btn = (Button)sender;
            ItemContainerGenerator gen;
            TreeNode node = FindViewNodeByTag(DirTreeView.ItemContainerGenerator, (string)btn.Tag, out gen);

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

            var tag = (string)btn.Tag;
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

            var tag = (string)btn.Tag;
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

        /// <summary>
        /// 从data集合里直接找
        /// </summary>
        /// <param name="tag"></param>
        /// <returns></returns>
        private TreeNode FindNodeInDataTree(string tag)
        {
            TreeNode retNode = null;
            bool inSearch = !string.IsNullOrEmpty(SearchBox.Text);
            if (inSearch)
            {
                retNode = _searchTreeNode.FindNodeInChild(tag);
            }

            if (retNode == null)
            {
                retNode = _rootNode.FindNodeInChild(tag);
            }

            return retNode;
        }

        /// <summary>
        /// 从ui绑定的data集合里取node
        /// </summary>
        /// <param name="tag"></param>
        /// <returns></returns>
        private TreeNode FindTreeViewNode(string tag)
        {
            ItemContainerGenerator gen;
            return FindViewNodeByTag(DirTreeView.ItemContainerGenerator, tag, out gen);
        }

        private TreeNode FindViewNodeByTag(ItemContainerGenerator container, string path, out ItemContainerGenerator generator)
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

                    var findNode = FindViewNodeByTag(treeViewItem.ItemContainerGenerator, path, out generator);
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

        private void HidePopupClick(object sender, RoutedEventArgs e)
        {
            SearchBinPopup.IsOpen = false;
        }

        private void CopyExcelName(object sender, RoutedEventArgs e)
        {
            var btn = (Button)sender;
            var excelName = btn.Tag;
            Clipboard.SetDataObject(excelName);

            HidePopupClick(null, null);
            SearchBox.Text = (string)excelName;
        }

        private void OpenSearchBinPopup(object sender, RoutedEventArgs e)
        {
            SearchBinPopup.IsOpen = true;
        }

        private void SearchBinTextChange(object sender, TextChangedEventArgs e)
        {
            InitBinName();
            var text = SearchBinBox.Text;
            List<BinListNode> resultList = new List<BinListNode>();
            SearchBin(text, resultList);
            BinResultList.ItemsSource = resultList;
        }

        private DateTime _lastParseTime;
        private List<BinListNode> _cacheBinList = new List<BinListNode>();
        private TreeNode _searchTreeNode;

        private void SearchBin(string searchBinName, List<BinListNode> resultList)
        {
            for (int i = 0; i < _cacheBinList.Count; i++)
            {
                var node = _cacheBinList[i];
                if (!string.IsNullOrEmpty(searchBinName) && Regex.Match(node.BinName, searchBinName, RegexOptions.IgnoreCase).Success)
                {
                    resultList.Add(node);
                }
            }
        }

        private void InitBinName()
        {
            Utils.ParseBinList(_cacheBinList, ref _lastParseTime);
        }

        private void OnSettingClick(object sender, RoutedEventArgs e)
        {
            SvnInfoWindow window = new SvnInfoWindow();
            window.Owner = this;
            window.Show();
        }

        private void OnCheckClick(object sender, RoutedEventArgs e)
        {
            ConvertDialog dlg = new ConvertDialog();
            dlg.Owner = this;
            dlg.OnClosedEvent += Utils.CleanConvert;
            dlg.Show();

            ServerLogChecker.ParseServerLog();
            if (ServerLogChecker.errorBinName != "")
                ServerLogChecker.AddErrorBin(ServerLogChecker.binExcelMap[ServerLogChecker.errorBinName], ref _rootNode);
        }

        private void ScanProgressBar_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {

        }

        private void InitComboBox()
        {
            var dirList = DirFilter.GetSelectDir();

            foreach (var dir in dirList)
            {
                comboBox.Items.Add(dir);
            }

            comboBox.SelectedIndex = comboBox.Items.IndexOf(DirFilter.defaultDir);
        }

        private void OnDirSelect(object sender, SelectionChangedEventArgs e)
        {
            DirFilter.SetSelectDir(comboBox.SelectedValue.ToString());
        }
    }
}
