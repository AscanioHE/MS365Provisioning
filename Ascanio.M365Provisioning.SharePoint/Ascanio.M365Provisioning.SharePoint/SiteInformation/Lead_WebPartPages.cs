using Ascanio.M365Provisioning.SharePoint.Services;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using System.Net.Http.Headers;
using File = Microsoft.SharePoint.Client.File;
using List = Microsoft.SharePoint.Client.List;

namespace Ascanio.M365Provisioning.SharePoint.SiteInformation
{
    public class Lead_WebPartPages
    {
        public Lead_WebPartPages()
        {
            using  (ClientContext context = new SharePointService().GetClientContext())
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
                        File pageFile = sitePage.File;
                        context.Load(pageFile);

                        context.Load(lib.RootFolder);
                        context.ExecuteQuery();

                        Console.WriteLine($"Processing page: {pageFile.ServerRelativeUrl}");

                        List pageList = pageFile.ListItemAllFields.ParentList;
                        string siteUrl = "https://23m2yz.sharepoint.com/sites/TestSite1";
                        string endpointUrl = $"{siteUrl}/_api/web/GetFileByServerRelativeUrl('{pageFile.ServerRelativeUrl}')/GetLimitedWebPartManager";
                        Uri endpointUri = new (endpointUrl, UriKind.Absolute);
                        _=GetWebParts(context, endpointUri);
                        string accessToken = context.GetAccessToken();
                        using (HttpClient client = new HttpClient())
                        {
                            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                            HttpResponseMessage response = await client.GetAsync(endpointUri);
                            if (response.IsSuccessStatusCode)
                            {
                                string result = await response.Content.ReadAsStringAsync();
                                // Analyseer de resultaten om informatie over de webonderdelen te verkrijgen
                                //...
                            }
                        }
                        //LimitedWebPartManager webPartManager = pageFile.GetLimitedWebPartManager(PersonalizationScope.Shared);
                        //context.Load
                        //    (
                        //    webPartManager.WebParts,
                        //    wp => wp.Include(wp => wp.WebPart.Title)
                        //    );
                        //context.ExecuteQuery();
                        //Console.WriteLine($"Number of web parts on the page: {webPartManager.WebParts.Count}");
                        //foreach (WebPartDefinition webPartDefinition in webPartManager.WebParts)
                        //{
                        //    Console.WriteLine($"WebPart Title : {webPartDefinition.WebPart.Title}");
                        //}
                    }
                }
            }

            Console.WriteLine("The SharePoint connection is closed");
        }

        private async Task GetWebParts(ClientContext context, Uri endpointUri)
        {
            string accessToken = context.GetAccessToken();
            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                HttpResponseMessage response = await client.GetAsync(endpointUri);
                if (response.IsSuccessStatusCode)
                {
                    string result = await response.Content.ReadAsStringAsync();
                    // Analyseer de resultaten om informatie over de webonderdelen te verkrijgen
                    //...
                }
            }
        }
    }
}
    