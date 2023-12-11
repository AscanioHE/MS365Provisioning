using M365Provisioning.SharePoint;
using M365Provisioning.SharePoint.DTO;
using Newtonsoft.Json;
using File = System.IO.File;

namespace M365Provisioning.SharePoint
{
    public class WriteData2Json
    {
        public void Write2JsonFile(object dtoFile, string jsonFilePath)
        {
            try
            {
                string json = JsonConvert.SerializeObject(dtoFile, Formatting.Indented);
                File.WriteAllText(jsonFilePath, json + Environment.NewLine);
            }
            catch (Exception ex)
            {
                // Log or print the exception details for debugging
                Console.WriteLine($"Error serializing WebTemplate: {ex.Message}");
            }
        }

        public string ConvertToJson(List<SiteSettingsDto> webTemplatesDTO)
        {
            string json = JsonConvert.SerializeObject(webTemplatesDTO);
            return json;
        }
    }
}
