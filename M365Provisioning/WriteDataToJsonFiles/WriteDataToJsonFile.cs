using System.Diagnostics;
using System.Dynamic;
using Newtonsoft.Json;
using File = System.IO.File;

namespace WriteDataToJsonFiles
{
    public class WriteDataToJsonFile : IWriteDataToJson
    {
        
        public object DtoFile = new();
        public string JsonFilePath { get; set; }
        public string JsonString { get; set; }

        public WriteDataToJsonFile()
        {
            JsonString = ConvertDtoToString();
        }

        public string ConvertDtoToString()
        {
            string jsonString = JsonConvert.SerializeObject(DtoFile, Formatting.Indented);
            
            return jsonString;
        }
        public string Write2JsonFile()
        {
            try
            {
                string json = JsonConvert.SerializeObject(DtoFile, Formatting.Indented);
                File.WriteAllText(JsonFilePath, json + Environment.NewLine);
                return json;
            }
            catch (Exception ex)
            {
                // Log or print the exception details for debugging
                Debug.WriteLine($"Error serializing WebTemplate : {ex.Message}");
            }
        }
    }
}