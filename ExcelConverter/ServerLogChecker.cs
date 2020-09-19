using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelConverter
{
    class ServerLogChecker
    {
        public static string ServerLogPath = "";
        public static string ServerLogPathDefine = "path_define.bat";
        public static string ServerLogPathKeyWord = "server_dir";

        public static void InitServerLogPath()
        {
            var PathDefine = File.ReadAllLines(Utils.WorkingPath + "\\" + ServerLogPathDefine);

            foreach (var line in PathDefine)
            {
                if (line.Contains(ServerLogPathKeyWord))
                {
                    Utils.DebugLog(line);
                    var equalIndex = line.IndexOf("=");
                    ServerLogPath = line.Substring(equalIndex + 1, line.Length - equalIndex - 1);
                    Utils.DebugLog(ServerLogPath);
                    break;
                }
            }
        }

        public static void ParseServerLog()
        {
            Utils.DebugLog(ServerLogPath);
        }
    }
}
