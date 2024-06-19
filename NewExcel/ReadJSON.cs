using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace JsontoExcel
{
    internal class ReadJSON
    {
        public ReadJSON()
        {
        }
        public static dynamic ReadJsonFile()
        {
            string directoryPath = @"C:\Train Simulator\Data\penilaian\";
            Directory.CreateDirectory(directoryPath);
            string[] jsonFiles = Directory.GetFiles(directoryPath, "*.json")
                                           .OrderByDescending(f => new FileInfo(f).LastWriteTime)
                                           .ToArray();

            if (jsonFiles.Length > 0)
            {
                string latestJsonFile = jsonFiles[0]; // Get the latest JSON file based on last write time
                dynamic jsonFile = JsonConvert.DeserializeObject(File.ReadAllText(latestJsonFile));
                return jsonFile;
            }
            else
            {
                Console.WriteLine("No JSON files found in the specified directory.");
                return null;
            }
        }
    }
}
