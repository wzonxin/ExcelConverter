using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Shapes;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;


namespace ExcelConverter
{
    class DirFilter
    {
        public static String defaultDir = "运营";
        public static String filterDirPatten = @"运营";
        public static int selectDirIndex = 0;
        public static List<string> filterDirs = new List<string>();

        public static List<string> GetSelectDir()
        {
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

        public static void SetSelectDir(String selectDir)
        {
            selectDirIndex = filterDirs.FindIndex(dir => selectDir == dir);
        }

        public static bool IsSkipDir(String selectDir)
        {
            bool bIsSkip = false;

            foreach (var skipDir in filterDirs)
            {
                if (skipDir == filterDirs[selectDirIndex])
                {
                    continue;
                }

                if (selectDir.Contains(skipDir + @"\"))
                {
                    return true;
                }
            }

            return bIsSkip;
        }
    }
}
