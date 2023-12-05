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
                List<List> list = GetListsWithWebParts(context);
            }

            Console.WriteLine("The SharePoint connection is closed");
        }

        private List<List> GetListsWithWebParts(ClientContext context)
        {
            ListCollection lists = context.Web.Lists;
            List<List> listsWithWebParts = new();
            context.Load
                (
                lists,
                l => l.Include
                              (
                              l => l.Title,
                              l => l.RootFolder,
                              l => l.RootFolder.ServerRelativeUrl,
                              l => l.DefaultViewUrl,
                              l => l.BaseTemplate,
                              l => l.RootFolder.Properties
                              )
                );
            context.ExecuteQuery();
            foreach (List list in lists)
            {
                context.Load
                    (
                    list.RootFolder,
                    f => f.Properties
                    );
                context.ExecuteQuery();
                if (list.RootFolder.Properties.FieldValues.ContainsKey("vti_pagecustomized") &&
                    (bool)list.RootFolder.Properties.FieldValues["vti_pagecustimized"])
                {
                    listsWithWebParts.Add(list);
                }
            }
            return listsWithWebParts;
        }
    }
}
