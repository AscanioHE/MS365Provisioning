using System.Diagnostics;
using System.Dynamic;
using System.Reflection;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.Json;
using Newtonsoft.Json;
using File = System.IO.File;

namespace WriteDataToJsonFiles
{
    public class WriteDataToJsonFile : IWriteDataToJson
    {
        public WriteDataToJsonFile()
        {
            string appSettingsPath = "appsettings.json";
            IConfigurationRoot configuration = new ConfigurationBuilder()
                .AddJsonFile(appSettingsPath, optional: false, reloadOnChange: true)
                .Build();
            WorkingDirectory = configuration["SharePoint:WorkingDirectoryPath"]!;
            Directory.SetCurrentDirectory(WorkingDirectory);
        }

        public object DtoFile { get; set; } = new();
        public string JsonFilePath { get; set; } = "TempJsonFile";
        public string WorkingDirectory { get; set; }

        public string ConvertDtoToString()
        {
            try
            {
                string jsonString = JsonConvert.SerializeObject(DtoFile, Formatting.Indented);
                return jsonString;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error creating Json String : {ex.Message}");
                throw;
            }
        }
        public string Write2JsonFile()
        {
            try
            {
                string jsonFile = JsonFilePath;
                JsonFilePath += $"{jsonFile}";
                string json = JsonConvert.SerializeObject(DtoFile, Formatting.Indented);
                File.WriteAllText(jsonFile, json + Environment.NewLine);
                return json;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error Writing Json String to file : {ex.Message}");
                throw;
            }
        }


    }
}