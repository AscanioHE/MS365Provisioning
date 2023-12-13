using M365Provisioning.SharePoint.Interfaces;
using M365Provisioning.SharePoint.Services;
using M365Provisioning.SharePoint.DTO;

using Microsoft.SharePoint.Client;

namespace M365Provisioning.SharePoint.Functions
{
    public class SharePointFunctions : ISharePointFunctions
    {


        private SharePointServices SharePointServices { get; set; }

        public List<SiteSettingsDto> Load()
        {
            SharePointServices = new SharePointServices();
            List<SiteSettingsDto> webTemplatesDto = new();
            ClientContext context = SharePointServices.Context;
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

            return webTemplatesDto;
        }

        public List<SiteSettingsDto> GetSiteSettings(ClientContext context)
        {
            List<SiteSettingsDto> siteSettings = new();

            return siteSettings;
        }

    }
}
