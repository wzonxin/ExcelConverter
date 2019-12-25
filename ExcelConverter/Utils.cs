using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Json;

namespace ExcelConverter
{
    public class Utils
    {
        private const string fileName = "fav.json";
        public static void SaveFav(List<TreeNode> pathList)
        {
            string runningPath = AppDomain.CurrentDomain.BaseDirectory;

            var jsonStr = JsonSerializer.Serialize(pathList);
            FileStream fileStream = File.Create(runningPath + fileName);
            fileStream.Write(Encoding.UTF8.GetBytes(jsonStr));
            fileStream.Flush(true);
            fileStream.Close();
        }

        public static List<TreeNode> ReadFav()
        {
            string runningPath = AppDomain.CurrentDomain.BaseDirectory;
            string filePath = runningPath + fileName;
            List<TreeNode> list = null;
            try
            {
                var str = File.ReadAllText(filePath);
                list = JsonSerializer.Deserialize<List<TreeNode>>(str);
            }
            catch (Exception e)
            {
                if (list == null)
                    list = new List<TreeNode>();
            }
            return list;
        }

        public static string GetFileName()
        {
            return "";
        }
    }
}
