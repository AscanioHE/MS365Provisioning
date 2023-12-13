using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace M365Provisioning.SharePoint.DTO
{
    public class SiteSettingsDto
    {
        public string SiteTemplate { get; set; } = string.Empty;
        public uint Value { get; set; }
        public SiteSettingsDto() { }
        public SiteSettingsDto(string siteTemplate, uint value)
        {
            SiteTemplate = siteTemplate;
            Value = value;
        }
    }
}
