using Newtonsoft.Json;

namespace Ascanio.M365Provisioning.SharePoint.SiteInformation
{
    public class SiteSettingsDTO
    {
        public string SiteTemplate { get; set; } = string.Empty;
        public uint Value { get; set; }
        public SiteSettingsDTO() { }
        public SiteSettingsDTO(string siteTemplate, uint value,)
        {
            SiteTemplate = siteTemplate;
            Value = value;
        }
    }
}
