using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SPList = Microsoft.SharePoint.Client.List;
using ListTemplateType = Microsoft.SharePoint.Client.ListTemplateType;
using Ascanio.M365Provisioning.SharePoint.SiteInformation;

namespace Ascanio.M365Provisioning.SharePoint.Services
{
    public class SharePointFunction : SharePointFunctionBase
    {
        public SharePointFunction()
        {
            GetAllSharePointItems();
        }
    }
    public class SharePointFunctionBase
    {
        public void GetAllSharePointItems()
        {
            SharePointService sharePointService = new();
            ClientContext context = sharePointService.GetClientContext();
            Web web = context.Web;
            Lead_SiteSettings lead_SiteSettings = new();
            //lead_SiteSettings.GetWebItemParameters(context, web);
            //_ = new Lead_Lists(context, web);
            //_ = new ListViews(context, web);
            _ =new Lead_WebPartPages(context, web);
            Console.WriteLine("Json files are created");
        }
    }

}

