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
            
            _ = new Lead_SiteSettings();
            _ = new Lead_Lists();
            //_ = new ListViews(context, web);
            //_ =new Lead_WebPartPages(context, web);
            _ = new Lead_FolderStructure();
            Console.WriteLine("Json files are created");
            //temp temp = new();
            //temp.Test();
        }
    }

}

