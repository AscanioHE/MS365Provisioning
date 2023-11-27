using Ascanio.M365Provisioning.SharePoint.Services;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using PnP.Core;

namespace Ascanio.M365Provisioning.SharePoint.SiteInformation
{
    public class Lead_SiteSettings
    {
         public void Main()
         {
            GetWebItemParameters();

         }

         private void GetWebItemParameters()
         {
            SharePointService sharePointService = new();
            ClientContext context = sharePointService.GetClientContext();
            Web web = context.Web;

            // Explicitly load the necessary properties
            context.Load(web, w => w.WebTemplate, w => w.Title);
            WebTemplateCollection webtTemplateCollection = web.GetAvailableWebTemplates(1033, true);
            context.Load(webtTemplateCollection);
            context.ExecuteQuery();
            foreach (WebTemplate webTemplate in webtTemplateCollection)
            {
                Console.WriteLine("Template Name: " + webTemplate.Name + "|    |" + webTemplate.Id + "|  |" + webTemplate.Lcid);
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

        }
    }
}

