using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MS365Provisioning.SharePoint.Model
{
    public class SharePointSettings
    {
        public string ClientId { get; set; }
        public string TenantId { get; set; }
        public string ThumbPrint { get; set; }
        public string SiteUrl { get; set; }

        public SharePointSettings()
        {
            
        }
    }
}
