using Newtonsoft.Json;

namespace M365Provisioning.SharePoint.Services
{
    public class WriteDataToJson
    {
        public void Write2JsonFile(object dtoFile, string jsonFilePath)
        {
            try
            {
                string json = JsonConvert.SerializeObject(dtoFile, Formatting.Indented);
                System.IO.File.WriteAllText(jsonFilePath, json + Environment.NewLine);
            }
            catch (Exception ex)
            {
                // Log or print the exception details for debugging
                Console.WriteLine($"Error serializing WebTemplate: {ex.Message}");
            }
        }

        public string ConvertToJsonString(object dtoFile)
        {
            string json = JsonConvert.SerializeObject(dtoFile, Formatting.Indented);
            return json;
        }
    }
}
