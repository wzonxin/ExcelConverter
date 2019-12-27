using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Json;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelConverter
{
    public class Utils
    {
        private const string favFileName = "fav.json";
        private const string treeFileName = "tree.json";
        private static string[] batFileContent;

        public static string WorkingPath = "";
        public static void InitWorkingPath()
        {
            //string runningPath = AppDomain.CurrentDomain.BaseDirectory;
            //var info = Directory.GetParent(System.Reflection.Assembly.GetExecutingAssembly().Location);
            //string path = info.FullName;

            string path = "";
#if DEBUG
            path = "C:\\Work\\data";
#else
            path = Environment.CurrentDirectory;
#endif
            Utils.WorkingPath = path;
        }

        public static void GenFileTree()
        {
            System.Threading.ThreadPool.QueueUserWorkItem(new System.Threading.WaitCallback(DoInBackgroundThread));
        }

        private static void DoInBackgroundThread(object state)
        {
            var xlsPath = Utils.WorkingPath + "\\xls\\";
            if (!Directory.Exists(xlsPath))
            {
                EventDispatcher.SendEvent(TaskType.SearchError, "xls文件夹不存在");
                return;
            }

            TreeNode rootNode = new TreeNode();
            Search(rootNode, xlsPath, NodeType.Dir);
            SaveFileTree(rootNode);
            EventDispatcher.SendEvent(TaskType.FinshedSearch, rootNode);
            GC.Collect();
        }

        private static void Search(TreeNode root, string path, NodeType nodeType)
        {
            root.Path = path;
            root.Name = Utils.GetTreeItemName(path, nodeType);

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

                EventDispatcher.SendEvent(TaskType.UpdateSearchProgress, (i + 1f) / childDirPath.Length);
            }

            for (int i = 0; i < files.Length; i++)
            {
                TreeNode node = new TreeNode();
                if (files[i].Contains("~$"))
                    continue;

                Search(node, files[i], NodeType.File);
                childNodes.Add(node);
            }

            root.Child = childNodes;
        }

        public static void FilterTree(TreeNode node, string filterStr, ref TreeNode filterNode)
        {
            filterNode = CloneTree(node);
            FilterTree(ref filterNode, filterStr);
            GC.Collect();
        }

        private static bool FilterTree(ref TreeNode node, string filterStr)
        {
            node.IsExpanded = true;
            var childs = node.Child;
            if (childs == null)
                return false;

            if (node.Name.Contains(filterStr, StringComparison.OrdinalIgnoreCase))
            {
                node.IsMatch = true;
                FindMatchFolderMatchFile(node, filterStr);
                return true;
            }

            bool bFind = false;
            for (int i = 0; i < childs.Count; i++)
            {
                TreeNode child = childs[i];
                if (child.IsFile)
                {
                    if (!childs[i].Name.Contains(filterStr, StringComparison.OrdinalIgnoreCase))
                    {
                        childs.RemoveAt(i);
                        i--;
                    }
                    else
                    {
                        child.IsMatch = true;
                        bFind = true;
                    }
                }
                else
                {
                    bool res = FilterTree(ref child, filterStr);
                    if (!res)
                    {
                        childs.RemoveAt(i);
                        i--;
                    }
                    else
                    {
                        bFind = true;
                    }
                }
            }
            return bFind;
        }

        //
        private static void FindMatchFolderMatchFile(TreeNode matchNode, string filterStr)
        {
            if (!matchNode.IsFile && matchNode.Child != null)
            {
                for (int i = 0; i < matchNode.Child.Count; i++)
                {
                    TreeNode childNode = matchNode.Child[i];
                    if (childNode.Name.Contains(filterStr, StringComparison.OrdinalIgnoreCase))
                    {
                        childNode.IsMatch = true;
                    }
                    FindMatchFolderMatchFile(childNode, filterStr);
                }
            }
        }

        private static TreeNode CloneTree(TreeNode tree)
        {
            TreeNode cloneNode = tree.Clone();
            if (tree.Child != null)
            {
                for (int i = 0; i < tree.Child.Count; i++)
                {
                    cloneNode.Child[i] = CloneTree(tree.Child[i]);
                }
            }
            return cloneNode;
        }

        private static void SaveFileTree(TreeNode treeNode)
        {
            var options = new JsonSerializerOptions
            {
                WriteIndented = true
            };
            var jsonStr = JsonSerializer.Serialize(treeNode, options);
            FileStream fileStream = File.Create(WorkingPath + "\\" + treeFileName);
            fileStream.Write(Encoding.UTF8.GetBytes(jsonStr));
            fileStream.Flush(true);
            fileStream.Close();
        }

        public static void SaveFav(List<TreeNode> pathList)
        {
            var options = new JsonSerializerOptions
            {
                WriteIndented = true
            };
            var jsonStr = JsonSerializer.Serialize(pathList, options);
            FileStream fileStream = File.Create(WorkingPath + "\\" + favFileName);
            fileStream.Write(Encoding.UTF8.GetBytes(jsonStr));
            fileStream.Flush(true);
            fileStream.Close();
        }

        public static TreeNode ReadTree()
        {
            string filePath = WorkingPath + "\\" + treeFileName;
            TreeNode root = null;
            try
            {
                var str = File.ReadAllText(filePath);
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true
                };
                root = JsonSerializer.Deserialize<TreeNode>(str, options);
            }
            catch (Exception e)
            {
                if (root == null)
                    root = new TreeNode();
            }
            return root;
        }
        
        public static List<TreeNode> ReadFav()
        {
            string filePath = WorkingPath + "\\" + favFileName;
            List<TreeNode> list = null;
            try
            {
                var str = File.ReadAllText(filePath);
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true
                };
                list = JsonSerializer.Deserialize<List<TreeNode>>(str, options);
            }
            catch (Exception e)
            {
                if (list == null)
                    list = new List<TreeNode>();
            }
            return list;
        }

        public static void ConvertExcel(List<TreeNode> convertList)
        {
            List<string> pathList = new List<string>();
            ConvertToPath(convertList, ref pathList);
            CopyXlsToTmpDir(ref pathList); 

            List<string> cmdList = new List<string>();
            for (int i = 0; i < pathList.Count; i++)
            {
                string filePath = pathList[i];
                //Create an object of FileInputStream class to read excel file
                FileInfo inputStream = new FileInfo(filePath);
                IWorkbook workbook = new XSSFWorkbook(inputStream);
                int sheetCnt = workbook.NumberOfSheets;
                for (int j = 0; j < sheetCnt; j++)
                {
                    ISheet sheet = workbook.GetSheetAt(j);
                    var sheetName = sheet.SheetName;
                    var batContent = GetBatCmd(sheetName);
                    if (!string.IsNullOrEmpty(batContent) && !cmdList.Contains(batContent))
                    {
                        cmdList.Add(batContent);
                    }
                }
                workbook.Close();
            }

            ExecuteCovertBat(cmdList);
            batFileContent = null;
        }

        private static void CopyXlsToTmpDir(ref List<string> pathList)
        {
            string copyStr = GetEnterDirStr() + @"
rd /S /Q .\xls_tmp
rd /S /Q .\csv
md .\xls_tmp
md .\csv
";
            for (int i = 0; i < pathList.Count; i++)
            {
                copyStr += "copy /y " + pathList[i] + " " + WorkingPath + "\\xls_tmp\\" + GetFileName(pathList[i]) + "\r\n";
                var fileName = GetFileName(pathList[i]);
                pathList[i] = pathList[i].Remove(pathList[i].LastIndexOf("\\xls\\")) + "\\xls_tmp\\" + fileName;
            }

            var batPath = CreateTmpBat(copyStr);
            var process = System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(batPath)
            {
                UseShellExecute = false,
                CreateNoWindow = true,
            });
            process.WaitForExit();
        }

        private static void ExecuteCovertBat(List<string> cmdList)
        {
            string prefix = @"set path=C:\Windows\System32;%path%" +
                            GetEnterDirStr()
+ 
@"rd /S /Q .\bin
rd /S /Q .\bin_cli
md .\bin
md .\bin_cli
call path_define.bat

del  build_err.log
del  build_info.log

";
            string middle = WorkingPath + "\\x2c\\xls2csv " + (WorkingPath + "\\xls_tmp\\ ") + (WorkingPath + "\\csv \"x2c.x2c\"\r\n\r\n");

            string cmd = "";
            for (int i = 0; i < cmdList.Count; i++)
            {
                cmd += cmdList[i] + "\r\n";
            }

            string last = "\r\npause";

            string batPath = CreateTmpBat(prefix + middle + cmd + last);
            var process = System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(batPath)
            {
                UseShellExecute = true,
            });
            process.WaitForExit();
        }

        private static string CreateTmpBat(string content)
        {
            string batPath = WorkingPath + "\\tmp.bat";
            var file = File.Create(batPath);
            file.Write(Encoding.GetEncoding("GBK").GetBytes(content));
            file.Flush(true);
            file.Close();
            return batPath;
        }

        private static string GetBatCmd(string sheetName)
        {
            if (batFileContent == null)
            {
                batFileContent = File.ReadAllLines(WorkingPath + "\\策划转表_公共.bat", Encoding.GetEncoding("GBK"));
            }

            for (var i = 0; i < batFileContent.Length; i++)
            {
                if (batFileContent[i].Contains(sheetName))
                    return batFileContent[i];
            }

            return "";
        }

        private static void ConvertToPath(List<TreeNode> nodes, ref List<string> pathList)
        {
            if (nodes.Count <= 0) return;

            for (var i = 0; i < nodes.Count; i++)
            {
                if (nodes[i].Type == NodeType.Dir)
                {
                    ConvertToPath(nodes[i].Child, ref pathList);
                }
                else
                {
                    if(!nodes[i].Path.Contains("~$"))
                        pathList.Add(nodes[i].Path);
                }
            }
        }

        private static string GetEnterDirStr()
        {
            var pathArr = WorkingPath.Split(":");
            var disk = pathArr[0];
            var folderPath = pathArr[1];
            return "\r\n" + disk + ":\r\n"
                + "cd " + folderPath + "\r\n";

            return @"
c:
cd \Work\data\\
";
        }

        public static string GetTreeItemName(string fullPath, NodeType nodeType)
        {
            var fileName = GetFileName(fullPath);
            if(nodeType == NodeType.File)
            {
                string subNames = GetSubSheetNames(fullPath);
                fileName += subNames;
            }
            return fileName;
        }

        public static string GetFileName(string fullPath)
        {
            var lastIndexOf = fullPath.LastIndexOf("\\", StringComparison.Ordinal) + 1;
            var name = fullPath.Substring(lastIndexOf, fullPath.Length - lastIndexOf);
            return name;
        }

        public static string GetSubSheetNames(string xlsPath)
        {
            FileInfo inputStream = new FileInfo(xlsPath);
            IWorkbook workbook = new XSSFWorkbook(inputStream);
            int sheetCnt = workbook.NumberOfSheets;
            string sheetNames = "(";
            for (int j = 0; j < sheetCnt; j++)
            {
                ISheet sheet = workbook.GetSheetAt(j);
                var sheetName = sheet.SheetName;
                sheetNames += sheetName;
                if (j < sheetCnt - 1)
                    sheetNames += ", ";
            }
            sheetNames += ")";
            return sheetNames;
        }

        public static int SortList(TreeNode node1, TreeNode node2)
        {
            if(node1.Type != node2.Type)
            {
                return (node1.Type - node2.Type);
            }
            else
            {
                return node1.SingleFileName.CompareTo(node2.SingleFileName);
            }
        }
    }
}
