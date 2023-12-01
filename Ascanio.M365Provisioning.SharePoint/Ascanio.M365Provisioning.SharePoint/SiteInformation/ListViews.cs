using Ascanio.M365Provisioning.SharePoint.Services;
using Microsoft.SharePoint.Client;
using PnP.Framework.Provisioning.Model.Configuration.Lists;
using System.Drawing.Text;

namespace Ascanio.M365Provisioning.SharePoint.SiteInformation
{
    public class ListViews
    {
        public ListViews(ClientContext context, Web web) 
        {
            web = context.Web;
            context.Load
                (
                web,
                w => w.Lists
                );
            context.ExecuteQuery();
            List<ListViewsDTO> listViewsDTO = GetListViews(context,web.Lists);
            string jsonFilePath = "JsonFiles/ListViews.json";
            WriteData2Json writeData2Json = new();
            writeData2Json.Write2JsonFile(listViewsDTO, jsonFilePath);
            context.Dispose();
        }  
        private List<ListViewsDTO> GetListViews(ClientContext context, ListCollection lists)
        {
            List<ListViewsDTO> listViewsDTO = new();
            foreach (List list in lists)
            {
                context.Load
                    (
                        list,
                        l => l.Views,
                        l => l.Title
                    );
                context.ExecuteQuery();
                foreach (View view in list.Views)
                {
                    context.Load
                        (
                            view,
                            v => v.Title,
                            v => v.DefaultView,
                            v => v.ViewFields,
                            v => v.RowLimit,
                            v => v.Scope
                        );
                    context.ExecuteQuery();
                    listViewsDTO.Add(new ListViewsDTO
                    {
                        ListName = list.Title,
                        ViewName = view.Title,
                        DefaultView = view.DefaultView,
                        ViewFields = GetViewFields(view),
                        RowLimit = view.RowLimit,
                        ListScope = view.Scope,
                        JsonFormatterFile = $"{list.Title}.json"
                    });

                }
            }
            return listViewsDTO;
        }
        private string GetViewFields (View view)
        {
            return string.Join(",",view.ViewFields);
        }
    }
}
