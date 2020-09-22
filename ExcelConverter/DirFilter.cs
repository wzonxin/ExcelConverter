using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Windows.Shapes;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;


namespace ExcelConverter
{
    class DirFilter
    {
        public static String defaultDir = "运营";
        public static String filterDirPatten = @"运营";
        public static Dictionary<String, Boolean> dirFilterMap;
        
        public static List<string> GetSelectDir()
        {
            List<string> filterDirs = new List<string>();

            DirectoryInfo di = new DirectoryInfo(@"xls");
            DirectoryInfo[] dirs = di.GetDirectories();

            foreach (DirectoryInfo dirInfo in dirs)
            {
                if (dirInfo.Name.Contains(filterDirPatten))
                {
                    filterDirs.Add(dirInfo.Name);
                }
            }

            return filterDirs;
        }
    }
}
