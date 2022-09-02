using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelConverter
{
    class ServerLogChecker
    {
        public static string serverPath = "";
        public static string serverLogPathDefine =  Utils.WorkingPath + "\\path_define.bat";
        public static string serverLogPathKeyWord = "server_dir";
        public static string serverLogPath = "";
        public static string serverLogPathFix = "\\log\\zone_svr";
        public static string serverLogPattern = "zone_svr_20*.log";

        public static string errorReadDataFailedKeyWord = "CBinReader.cpp:273";
        public static string errorReadDataFailedInfo = "协议不一致";
        public static string errorReadDataFailedPatten = @"cfg/res/(\w+).bin";

        public static Dictionary<String, List<String>> binExcelMap = new Dictionary<String, List<String>>();
        public static String errorBinName = new String("");

        public static void InitServerLogPath()
        {
            var PathDefine = File.ReadAllLines(serverLogPathDefine);

            foreach (var line in PathDefine)
            {
                if (line.Contains(serverLogPathKeyWord))
                {
                    var equalIndex = line.IndexOf("=");
                    serverPath = line.Substring(equalIndex + 1, line.Length - equalIndex - 1);
                    break;
                }
            }

            serverLogPath = serverPath + serverLogPathFix;
        }

        public static void InitBinExcelMap()
        {
            var strFlePath = Utils.WorkingPath + "\\策划转表_公共.bat";
            var binExcelMapFile = File.ReadAllLines(strFlePath, Encoding.GetEncoding("GBK"));

            foreach (var line in binExcelMapFile)
            {
                if (line.Contains("do_conv"))
                {
                    string[] words = line.Split(" ", options: StringSplitOptions.RemoveEmptyEntries);
                    List<string> excels = new List<string>();
                    for (int i = 3; i < words.Length; i++)
                    {
                        excels.Add(words[i]);
                    }

                    if (binExcelMap.ContainsKey(words[2]) == false)
                    {
                        binExcelMap.Add(words[2], excels);
                    }
                    else
                    {
                        binExcelMap[words[2]].AddRange(excels);
                    }
                }
            }
        }

        public static string GetLatestLogFileName()
        {
            DirectoryInfo di = new DirectoryInfo(serverLogPath);
            var logFiles = di.GetFiles(serverLogPattern);

            return logFiles[logFiles.Length - 1].Name;
        }

        private static string[] GetLastLines(FileStream fs, int n)
        {
            int seekLength = (int)(fs.Length < n ? fs.Length : n);
            byte[] buffer = new byte[seekLength];
            fs.Seek(-buffer.Length, SeekOrigin.End);
            fs.Read(buffer, 0, buffer.Length);
            string multLine = System.Text.Encoding.GetEncoding("gb2312").GetString(buffer);
            string[] lines = multLine.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
 
            return lines;
        }

        public static string GetErrorBinName(string patten, string line)
        {
            Regex rex = new Regex(patten, RegexOptions.IgnoreCase);
            MatchCollection matches = rex.Matches(line);
            GroupCollection groups = matches[0].Groups;
            return groups[1].Value;
        }

        public static void AddErrorBin(List<String> excelNames,ref TreeNode node)
        {
            if (!node.IsFile && node.Child != null)
            {
                for (int i = 0; i < node.Child.Count; i++)
                {
                    TreeNode childNode = node.Child[i];

                    foreach (var excelName in excelNames)
                    {
                        if (childNode.MatchSearch(excelName) && !DirFilter.IsSkipDir(childNode.Path))
                        {
                            Utils.DebugLog(childNode.Name);
                            Utils.DebugLog(childNode.Path);
                            childNode.IsOn = true;
                            break;
                        }
                    }

                    AddErrorBin(excelNames, ref childNode);
                }
            }
        }

        public static void ParseServerLog()
        {
            Utils.DebugLog(serverLogPath);
            var targetLogFile = GetLatestLogFileName();
            Utils.DebugLog(targetLogFile);

            var fileName = serverLogPath + "\\" + targetLogFile;

            FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

            // 读取末N行                    
            var lines = GetLastLines(fs, 10240);
            foreach (var Line in lines)
            {
                if (Line.Contains(errorReadDataFailedKeyWord))
                {
                    errorBinName = GetErrorBinName(errorReadDataFailedPatten, Line);
                    Utils.DebugLog(errorBinName + errorReadDataFailedInfo);
                    break;
                }
            }
        }
    }
}
