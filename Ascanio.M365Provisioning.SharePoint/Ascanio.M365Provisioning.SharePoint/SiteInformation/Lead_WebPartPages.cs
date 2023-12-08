using Ascanio.M365Provisioning.SharePoint.Services;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using Newtonsoft.Json;
using System.Net.Http.Headers;
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
                        sp => sp.Include(sp => sp.Client_Title)
                        );
                    context.ExecuteQuery();
                    //foreach (ListItem sitePage in sitePages)
                    //{
                    //    string fullPageUrl = "https://23m2yz.sharepoint.com/sites/TestSite1/SitePages/Home.aspx";

                    //    File pageFile = context.Web.GetFileByServerRelativeUrl(fullPageUrl);
                    //    context.Load(pageFile);
                    //    context.ExecuteQuery();

                    //    var page = ClientSidePage.Load(context, pageFile);
                    //    var components = page.Controls;

                    //    foreach (var component in components)
                    //    {
                    //        if (component is ClientSideWebPart)
                    //        {
                    //            var webPart = (ClientSideWebPart)component;
                    //            Console.WriteLine($"WebPart Title: {webPart.Title}");
                    //            //Console.WriteLine($"WebPart Id: {webPart.Id}");
                    //            // Voeg andere eigenschappen toe die je nodig hebt
                    //        }
                    //    }
                    //}
                }
            }

                Console.WriteLine("WebParts.json File created...");
            Console.WriteLine("The SharePoint connection is closed");
        }

    }
    public class WebPartInfo
    {
        public List<WebPart>? WebParts { get; set; } = null;
    }

    public class WebPart
    {
        public string Title { get; set; } = string.Empty;
        public string Type { get; set; } = string.Empty;
        // Voeg andere eigenschappen toe die je nodig hebt
    }
}
    