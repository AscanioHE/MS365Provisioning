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
            //_ = new Lead_WebPartPages();
            _ = new SiteSettings();
            Console.WriteLine("Lead_SiteSettings.json File created...");
            _ = new Lists();
            Console.WriteLine("Lead_Lists.json File created...");
            _ = new ListViews();
            Console.WriteLine("ListViews.json File created...");
            _ = new FolderStructure();
            Console.WriteLine("Lead_FolderStructure.json File created...");
            Console.WriteLine("Json files are created");
        }
    }
            //
            //temp temp = new();
            //temp.Test();

}

