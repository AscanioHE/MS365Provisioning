using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using PnP.Core.Model.SharePoint;
using System.IO;
using File = System.IO.File;

namespace Ascanio.M365Provisioning.SharePoint.Services
{
    public class WriteData2Json
    {
        public void Write2JsonFile(List<WebTemplate> webTemplates, string filePath)
        {
            foreach (WebTemplate webTemplate in webTemplates)
            {
                string json = JsonConvert.SerializeObject(webTemplate, Formatting.Indented);
                File.WriteAllText(filePath, json);
            }
        }
    }
}
