using Ascanio.M365Provisioning.SharePoint.Services;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using File = Microsoft.SharePoint.Client.File;
using List = Microsoft.SharePoint.Client.List;

namespace Ascanio.M365Provisioning.SharePoint.SiteInformation
{
    public class Lead_WebPartPages
    {
        public Lead_WebPartPages()
        {
            using (ClientContext context = new SharePointService().GetClientContext())
            {
                Console.WriteLine("Lead_SiteSettings.json File created...");
                IEnumerable<List> Libraries = context.LoadQuery
                                                            (
                                                            context.Web.Lists.Where
                                                                                (
                                                                                l => l.BaseTemplate == (int)ListTemplateType.WebPageLibrary
                                                                                )
                                                            );
                context.ExecuteQuery();
                foreach (List lib in Libraries)
                {
                    CamlQuery query = CamlQuery.CreateAllItemsQuery();
                    ListItemCollection sitePages = lib.GetItems(query);
                    context.Load
                        (
                        sitePages,
                        sp => sp.Include(sp=> sp.Client_Title)
                        );
                    context.ExecuteQuery();
                    foreach(ListItem sitePage in sitePages)
                    {
                        Console.WriteLine(sitePage.Client_Title);
                    }
                }
            }

            Console.WriteLine("The SharePoint connection is closed");
        }
    }
}
    