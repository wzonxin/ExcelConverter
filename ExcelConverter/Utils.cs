using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Text.Unicode;

namespace ExcelConverter
{
    public class Utils
    {
        public static void SaveFav(List<string> pathList)
        {
            string runningPath = AppDomain.CurrentDomain.BaseDirectory;

            var jsonStr = JsonSerializer.Serialize(pathList);
            FileStream fileStream = File.Create(runningPath + "fav.json");
            fileStream.Write(Encoding.UTF8.GetBytes(jsonStr));
        }
    }
}
