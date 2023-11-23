using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SPList = Microsoft.SharePoint.Client.List;
using ListTemplateType = Microsoft.SharePoint.Client.ListTemplateType;

namespace Ascanio.M365Provisioning.SharePoint.Services
{
    public class SharePointFunction
    {
        readonly SharePointService sharePointService = new();

        public SharePointFunction()
        {
            ClientContext client = sharePointService.GetClientContext();

        }
        
       
    }
}
