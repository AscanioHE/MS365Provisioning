using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SPList = Microsoft.SharePoint.Client.List;
using ListTemplateType = Microsoft.SharePoint.Client.ListTemplateType;
using Ascanio.M365Provisioning.SharePoint.SiteInformation;

namespace Ascanio.M365Provisioning.SharePoint.Services
{
    public class SharePointFunction
    {
        public SharePointFunction()
        {
            Lead_SiteSettings leadSiteSettings = new();
            leadSiteSettings.Main();

        }
        
       
    }
}

