using Microsoft.SharePoint.Client;
using Ascanio.M365Provisioning.SharePoint.Services;

namespace Ascanio.M365Provisioning.SharePoint.SiteInformation
{
    public class SiteColumns
    {
        public SiteColumns() 
        { 
            SharePointService sharePointService = new();
            using(ClientContext context = sharePointService.GetClientContext())
            {
                Web web = context.Web;
                context.Load(
                            web,
                            w => w.Fields,
                            w => w.Fields.Include(f => f.Hidden,
                                                  f => f.InternalName,
                                                  f => f.SchemaXml,
                                                  f => f.DefaultValue
                                                  )
                            );
                context.ExecuteQuery();

                foreach(Field siteColumn in web.Fields)
                {
                    if(!siteColumn.Hidden)
                    {
                        Console.WriteLine($"Name: {siteColumn.Title}");
                        Console.WriteLine($"SchemaXml: {siteColumn.SchemaXml}");
                        Console.WriteLine($"DefaultValue: {siteColumn.DefaultValue}");
                    }
                }
            }
        }

    }
}
