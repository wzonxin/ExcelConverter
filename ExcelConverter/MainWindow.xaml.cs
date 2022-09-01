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
    // ReSharper disable once UnusedMember.Global
    // ReSharper disable once RedundantExtendsListEntry
    public partial class MainWindow : Window
    {
        public List<TreeNode> DataSource { get; set; }
        private List<TreeNode> _convertList = new List<TreeNode>();
        private List<TreeNode> _favList;
        private TreeNode _rootNode;
        private System.Windows.Threading.DispatcherTimer _timer;
        private string m_lastSearchInput;

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
            _timer.Tick += TimerTick;
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
            EventDispatcher.RegdEvent(TaskType.ConvertFinishWithFailed, OnConvertFinishWithFailed);
        }

        private void TimerTick(object s, EventArgs a)
        {
            EventDispatcher.CheckTick();
            TickNodeChecked();
        }

        private void OnWindowClosed(object sender, EventArgs e)
        {
            _timer?.Stop();
            EventDispatcher.Clear();
            MemPoolMgr.Instance.ClearAllPool();
        }

        private void OnWindowDeactivated(object sender, EventArgs e)
        {
            HidePopupClick(null, null);
        }

        private void RefreshConvertList()
        {
            _convertList.Sort(Utils.SortList);

            float height = 30;

            ConvertGrid.Children.Clear();
            for (int i = 0; i < _convertList.Count; i++)
            {
                TreeNode convertNode = _convertList[i];
                if (!convertNode.IsFile)
                {
                    continue;
                }

                Button bt = GenConvertItem(height, convertNode);

                Button removeBt = new Button();
                removeBt.Tag = convertNode.Path;
                removeBt.Click += RemoveCovertBtnClick;
                removeBt.Content = "X";
                removeBt.Width = 20;
                removeBt.Height = height - 10;

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
            item.Header = "同步DR并转表";
            item.Click += (obj, ev) =>
            {
                DoConvert(new List<TreeNode>() { treeNode }, true);
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
            SetTreeSource(rootNode);
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

            for (var i = 0; i < list.Count; i++)
            {
                var node = FindNodeInDataTree(list[i]);
                if (node != null && !DirFilter.IsSkipDir(node.Path))
                {
                    //node.SetChecked(true);

                    if (!_convertList.Contains(node))
                    {
                        _convertList.Add(node);
                    }
                }
            }

            RefreshConvertList();
            SyncConvert2TreeShow();
        }

        private void DoConvert(List<TreeNode> nodeList, bool? upDr = null)
        {
            if (_dlg != null)
            {
                return;
            }

            ConvertDialog dlg = new ConvertDialog();
            dlg.Owner = this;
            dlg.OnClosedEvent += OnConvertDialogClose;
            _dlg = dlg;
            dlg.Show();

            Utils.ConvertExcel(nodeList, upDr);
        }

        private void OnConvertDialogClose()
        {
            _dlg = null;
            Utils.CleanConvert();
        }

        private void SearchTextChange(object sender, TextChangedEventArgs e)
        {
            try
            {
                string inputText = SearchBox.Text;
                if (inputText.Equals(m_lastSearchInput))
                {
                    return;
                }

                if (!string.IsNullOrEmpty(inputText))
                {
                    if (!string.IsNullOrEmpty(m_lastSearchInput) && inputText.Contains(m_lastSearchInput))
                    {
                        //直接从上次结果找
                        _searchTreeNode = Utils.SearchTree(_searchTreeNode, inputText);
                    }
                    else
                    {
                        _searchTreeNode = Utils.SearchTree(_rootNode, inputText);
                    }

                    SetTreeSource(_searchTreeNode);
                    SyncConvert2TreeShow();
                }
                else
                {
                    _searchTreeNode = null;
                    SetTreeSource(_rootNode);

                    SyncConvert2TreeShow();
                    UpdateParentNodeCheck(true);
                }
                FuncTabControl.SelectedIndex = 0;
                m_lastSearchInput = inputText;
            }
            catch (Exception exception)
            {
                string logContent = exception.Message + "\n" + Environment.StackTrace;
                MessageBox.Show("搜索有误，请查看日志excel_debug.log\n" + logContent);
                Utils.SaveLogFile(logContent);
            }
        }
        
        private void ClickClearSearchBtn(object sender, RoutedEventArgs e)
        {
            SearchBox.Text = string.Empty;
        }

        private void SetTreeSource(TreeNode node)
        {
            DataSource = node.Child;
            DirTreeView.ItemsSource = DataSource;

            EmptyTipPanel.Visibility = node.Child == null || node.Child.Count == 0 ? Visibility.Visible : Visibility.Hidden;
        }

        private void ScanDir(object sender, RoutedEventArgs e)
        {
            ScanLabel.Content = "扫描中...";
            Utils.GenFileTree();
        }

        private void OnClearConvertList(object sender, RoutedEventArgs e)
        {
            _convertList.Clear();
            RefreshConvertList();

            if (_searchTreeNode != null)
            {
                _searchTreeNode.SetChecked(false);
            }

            _rootNode.SetChecked(false);
        }
        
        private void UpdateProgress(float value)
        {
            float val = value;
            ScanProgressBar.Value = val * 100;
        }

        private void OnFinishedSearch(TreeNode node)
        {
            _rootNode = node;
            SetTreeSource(node);
            ScanLabel.Content = "扫描完成";
            SearchBox.Text = null;

            MessageBox.Show("扫描完成");
        }

        private void OnConvertFinishWithFailed()
        {
            var result = MessageBox.Show("转表报错，同步DR并重新转表？");
            if (result == MessageBoxResult.Yes)
            {
                if (_dlg != null)
                {
                    _dlg.Close();
                }
                DoConvert(_convertList);
            }
        }

        private void OnSearchError(string errorStr)
        {
            //ScanLabel.Content = errorStr;
            MessageBox.Show($"{errorStr}");
        }
        
        private void SyncConvert2TreeShow()
        {
            Dictionary<string, bool> dictNodeList = new Dictionary<string, bool>(_convertList.Count);
            foreach (var treeNode in _convertList)
            {
                dictNodeList[treeNode.Path] = true;
            }

            _rootNode.Recursive(t =>
            {
                t.IsOn = dictNodeList.ContainsKey(t.Path);
            });

            if (!string.IsNullOrEmpty(SearchBinBox.Text))
            {
                _searchTreeNode.Recursive(t =>
                {
                    t.IsOn = dictNodeList.ContainsKey(t.Path);
                });
            }
        }

        private void SyncTree2ConvertShow()
        {
            _convertList.Clear();

            _rootNode.Recursive(t =>
            {
                if (t.IsOn)
                {
                    _convertList.Add(t);
                }
            });
            
            bool inSearch = !string.IsNullOrEmpty(SearchBox.Text);
            if (inSearch)
            {
                _searchTreeNode.Recursive(t =>
                {
                    TreeNode fiNode;
                    if (!t.IsOn && (fiNode = _convertList.Find(lt => lt.Path == t.Path)) != null)
                    {
                        _convertList.Remove(fiNode);
                    }

                    if (t.IsOn && !_convertList.Contains(t))
                    {
                        _convertList.Add(t);
                    }
                });
            }

            RefreshConvertList();
        }

        private void MenuItemAddToFavClick(object sender, RoutedEventArgs e)
        {
            var tag = ((MenuItem)sender).Tag;
            var node = FindNodeInDataTree((string)tag);
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
            var node = FindNodeInDataTree((string)tag);

            DoConvert(new List<TreeNode>() { node });
        }

        private void TreeItemUpDrConvert(object sender, RoutedEventArgs e)
        {
            var tag = ((MenuItem)sender).Tag;
            var node = FindNodeInDataTree((string)tag);

            DoConvert(new List<TreeNode>() { node }, true);
        }

        //private void TreeItemAddConvert(object sender, RoutedEventArgs e)
        //{
        //    var tag = ((MenuItem)sender).Tag;
        //    var node = FindTreeViewNode((string)tag);

        //    //AddConvertNode(node);
        //}

        private void OpenFavItemFolder(object sender, RoutedEventArgs e)
        {
            var tag = ((MenuItem)sender).Tag;
            var node = _favList.Find(treeNode => (string)tag == treeNode.Path);
            if (node != null)
            {
                node.AutoOpen();
            }
        }

        private void OpenTreeItemFolder(object sender, RoutedEventArgs e)
        {
            var tag = ((MenuItem)sender).Tag;
            var node = FindNodeInDataTree((string)tag);
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
            var node = _favList.Find(treeNode => tag == treeNode.Path);
            if (node == null)
            {
                return;
            }

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
            var node = _convertList.Find(treeNode => tag == treeNode.Path);
            if (node == null)
            {
                return;
            }

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
            var treeViewItem = e.Source as TreeViewItem;
            if (treeViewItem == null)
            {
                return;
            }

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
        private ConvertDialog _dlg;

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
            if (_dlg != null)
            {
                return;
            }

            ConvertDialog dlg = new ConvertDialog();
            dlg.Owner = this;
            dlg.OnClosedEvent += OnConvertDialogClose;
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

        private void OnTreeViewItemToggleValueChanged(object sender, RoutedEventArgs e)
        {
            CheckBox checkBox = (CheckBox)sender;
            var tag = (string)checkBox.Tag;
            var treeNode = FindNodeInDataTree(tag);
            if (treeNode != null)
            {
                treeNode.SetChecked(checkBox.IsChecked != null && (bool)checkBox.IsChecked);

                //SyncTree2ConvertShow();

                _toggleChanged = true;
            }
        }

        private bool _toggleChanged;

        private void TickNodeChecked()
        {
            UpdateParentNodeCheck();
        }

        private void UpdateParentNodeCheck(bool forceRefresh = false)
        {
            if (!_toggleChanged && !forceRefresh)
            {
                return;
            }

            SyncTree2ConvertShow();

            TreeNode._banChildChange = true;

            _curTreeNode.Recursive(t =>
            {
                if (!t.IsFile)
                {
                    bool childAllOn = t.CheckChildAllOn();
                    if (t.IsOn != childAllOn)
                    {
                        t.SetChecked(childAllOn);
                    }
                }
            });

            TreeNode._banChildChange = false;
            _toggleChanged = false;
        }
    }
}
