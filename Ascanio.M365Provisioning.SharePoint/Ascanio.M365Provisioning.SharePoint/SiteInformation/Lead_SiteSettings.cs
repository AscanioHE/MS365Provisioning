using Ascanio.M365Provisioning.SharePoint.Services;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using PnP.Core;
using PnP.Core.Model.SharePoint;
using System.IO;
using File = System.IO.File;

namespace Ascanio.M365Provisioning.SharePoint.SiteInformation
{
    public class Lead_SiteSettings
    {
        public Lead_SiteSettings()
         {
            SharePointService sharePointService = new();
            ClientContext context = sharePointService.GetClientContext();
            Web web = context.Web;
            // Explicitly load the necessary properties
            context.Load(
                web,
                w => w.WebTemplate
                );
            context.ExecuteQuery();
            WebTemplateCollection webtTemplateCollection = web.GetAvailableWebTemplates(1033, true);
            context.Load(webtTemplateCollection);
            context.ExecuteQuery();

            List<Lead_SiteSettingsDTO> webTemplatesDTO = new();

            foreach (WebTemplate template in webtTemplateCollection)
            {
                // Create a Lead_SiteSettingsDTO and add it to the list
                webTemplatesDTO.Add(new Lead_SiteSettingsDTO
                {
                    SiteTemplate = template.Name,
                    Value = template.Lcid
                    // Other properties as needed
                });
            }
            try
            {
                // Execute the query to retrieve the data
                context.ExecuteQuery();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error executing query: {ex.Message}");
                // Handle the exception as needed
            }
            string jsonFilePath = "JsonFiles/Lead_SiteSettings.json";
            WriteData2Json writeData2Json = new();
            writeData2Json.Write2JsonFile(webTemplatesDTO,jsonFilePath);
            context.Dispose();
        }
    }
}

