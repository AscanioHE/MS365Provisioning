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
    public class SiteSettings : ISiteSettingsService
    {
        public SiteSettings()
        {
            //SharePointService sharePointService = new();
            //ClientContext context = sharePointService.GetClientContext();
            //Web web = context.Web;
            //context.Load(
            //                web
            //            );
            //context.ExecuteQuery();

            //WebTemplateCollection webtTemplateCollection = web.GetAvailableWebTemplates(1033, true);
            //context.Load(webtTemplateCollection);
            //context.ExecuteQuery();

            //List<SiteSettingsDTO> webTemplatesDTO = new();

            //foreach (WebTemplate template in webtTemplateCollection)
            //{
            //    if (!template.IsHidden)
            //    {
            //        // Create a Lead_SiteSettingsDTO and add it to the list
            //        webTemplatesDTO.Add(new SiteSettingsDTO
            //        {
            //            SiteTemplate = template.Name,
            //            Value = template.Lcid
            //            // Other properties as needed
            //        });
            //    }
            //}
            //try
            //{
            //    // Execute the query to retrieve the data
            //    context.ExecuteQuery();
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine($"Error executing query: {ex.Message}");
            //    // Handle the exception as needed
            //}
            //string jsonFilePath = sharePointService.SiteSettingsFilePath;
            //WriteData2Json writeData2Json = new();
            //writeData2Json.Write2JsonFile(webTemplatesDTO,jsonFilePath);
            //context.Dispose();
        }

        public List<SiteSettingsDTO> Load()
        {
            SharePointService sharePointService = new();
            ClientContext context = sharePointService.GetClientContext();
            Web web = context.Web;
            context.Load(
                web
            );
            context.ExecuteQuery();

            WebTemplateCollection webtTemplateCollection = web.GetAvailableWebTemplates(1033, true);
            context.Load(webtTemplateCollection);
            context.ExecuteQuery();

            List<SiteSettingsDTO> webTemplatesDTO = new();

            foreach (WebTemplate template in webtTemplateCollection)
            {
                if (!template.IsHidden)
                {
                    // Create a Lead_SiteSettingsDTO and add it to the list
                    webTemplatesDTO.Add(new SiteSettingsDTO
                    {
                        SiteTemplate = template.Name,
                        Value = template.Lcid
                        // Other properties as needed
                    });
                }
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
            finally
            {
                context.Dispose();
            }

            return webTemplatesDTO;
        }

        public string ConvertToJson(List<SiteSettingsDTO> webTemplatesDTO)
        {
            string json = JsonConvert.SerializeObject(webTemplatesDTO);
            return json;
        }

        public void WriteToJsonFile(List<SiteSettingsDTO> webTemplatesDTO, string jsonFilePath)
        {
            //string jsonFilePath = sharePointService.SiteSettingsFilePath;
            
            WriteData2Json writeData2Json = new();
            writeData2Json.Write2JsonFile(webTemplatesDTO, jsonFilePath);
        }
    }
}

