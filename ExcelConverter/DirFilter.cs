using System;
using System.Collections.Generic;
using System.IO;


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
            var xlsPath = Utils.WorkingPath + "\\xls";
            DirectoryInfo di = new DirectoryInfo(xlsPath);
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
