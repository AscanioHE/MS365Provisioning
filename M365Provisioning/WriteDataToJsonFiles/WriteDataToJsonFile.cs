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
            try
            {
                string jsonString = JsonConvert.SerializeObject(DtoFile, Formatting.Indented);
                return jsonString;

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating Json String : {ex.Message}");
                throw;
            }
        }
        public string Write2JsonFile()
        {
            try
            {
                string json = JsonConvert.SerializeObject(DtoFile, Formatting.Indented);
                try
                {
                    File.WriteAllText(JsonFilePath, json + Environment.NewLine);
                    return json;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error Writing Json String to file : {ex.Message}");
                    throw;
                }
            }
            catch (Exception ex)
            {
                // Log or print the exception details for debugging
                Debug.WriteLine($"Error creating Json String : {ex.Message}");
                return string.Empty;
            }
        }
    }
}