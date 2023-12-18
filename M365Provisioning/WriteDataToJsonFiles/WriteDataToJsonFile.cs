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

            public object DtoFile { get; set; } = new();
            public string JsonFilePath { get; set; } = "TempJsonFile";
        public WriteDataToJsonFile()
        {
            string appSettingsPath = "appsettings.json";
            IConfigurationRoot configuration = new ConfigurationBuilder()
                .AddJsonFile(appSettingsPath, optional: false, reloadOnChange: true)
                .Build();
        }

        public string ConvertDtoToString()
        {
            try
            {
                string jsonString = JsonConvert.SerializeObject(DtoFile, Formatting.Indented);
                return jsonString;
            }
            catch (Exception ex)
            {
                //_logger?.LogInformation($"Error creating Json String : {ex.Message}");
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
                //_logger?.LogInformation($"Error Writing Json String to file : {ex.Message}");
                throw;
            }
        }


    }
}