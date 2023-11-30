using Ascanio.M365Provisioning.SharePoint.SiteInformation;
using Newtonsoft.Json;
using File = System.IO.File;

namespace Ascanio.M365Provisioning.SharePoint.Services
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
    }

    

}
