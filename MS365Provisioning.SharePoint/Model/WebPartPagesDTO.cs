using Microsoft.Extensions.Logging;
using PnP.Core.Model.SharePoint;
using System.Net.Sockets;

namespace MS365Provisioning.SharePoint.Model
{
    public class WebPartPagesDto
    {
        public string Title { get; set; }
        public string Name { get; set; }
        public string QuickLaunchHeader { get; set; }
        public bool ShowComments { get; set; }
        public Type WebPartType { get; set; }
        public List<WebPartItem> webPartItems { get; set; }
        public string List { get; set; }
        public string View { get; set; }
        
        public WebPartPagesDto() 
        {
            Title = string.Empty;
            Name = string.Empty;
            QuickLaunchHeader = string.Empty;
            ShowComments = false;
            WebPartType = typeof(object);
            webPartItems = new List<WebPartItem>();
            List = string.Empty;
            View = string.Empty;
        }
    }
    public class WebPartItem
    {
        public string Name { get; set; }
        public string WebPartID { get; set; }
        public string PropertiesJson { get; set; }
        public WebPartItem()
        {
            Name=string.Empty;
            WebPartID=string.Empty;
            PropertiesJson=string.Empty;
        }
    }
}
