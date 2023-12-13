using System.Diagnostics.Contracts;
using M365Provisioning.SharePoint.Interfaces;
using M365Provisioning.SharePoint.Services;
using M365Provisioning.SharePoint.DTO;

using Microsoft.SharePoint.Client;

namespace M365Provisioning.SharePoint.Functions
{
    public class SharePointFunctions : ISharePointFunctions
    {


        private ISharePointServices SharePointServices { get; set; } = new SharePointServices();

        public List<SiteSettingsDto> LoadSiteSettings()
        {
            string jsonFilePath = SharePointServices.SiteSettingsFilePath;

            List<SiteSettingsDto> webTemplatesDto = new();
            ClientContext context = SharePointServices.GetClientContext();
            Web web = context.Web;
            context.Load(web);
            try
            {
                context.ExecuteQuery();

                WebTemplateCollection webtTemplateCollection = web.GetAvailableWebTemplates(1033, true);
                context.Load(webtTemplateCollection);
                context.ExecuteQuery();


                foreach (WebTemplate template in webtTemplateCollection)
                {
                        webTemplatesDto.Add(new SiteSettingsDto
                        {
                            SiteTemplate = template.Name,
                            Value = template.Lcid
                        });
                }
                
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error executing query: {ex.Message}");
                return new List<SiteSettingsDto>();
            }
            finally
            {
                context.Dispose();
            }
            WriteDataToJson writeDataToJson = new ();
            writeDataToJson.Write2JsonFile(webTemplatesDto, jsonFilePath);
            return webTemplatesDto;
        }

        public List<ListDto> GetLists()
        {
            List<ListDto> listDtos = new ();
            ClientContext context = SharePointServices.GetClientContext();
            ListCollection listCollection = context.Web.Lists;
            context.Load(listCollection,
                         lc => lc.Where(
                                                                    l =>l.Hidden == false));
            try
            {
                context.ExecuteQuery();

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }
            finally
            {
                context.Dispose();
            }
            return listDtos;
        }
    }
}
