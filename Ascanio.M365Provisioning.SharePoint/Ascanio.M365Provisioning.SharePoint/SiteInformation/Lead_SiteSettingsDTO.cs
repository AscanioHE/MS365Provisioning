using Newtonsoft.Json;

namespace Ascanio.M365Provisioning.SharePoint.SiteInformation
{
    public class Lead_SiteSettingsDTO
    {
        public string SiteTemplate { get; set; } = string.Empty;
        public uint Value { get; set; }

        public Lead_SiteSettingsDTO() { }



        public Lead_SiteSettingsDTO(string siteTemplate, uint value)
        {
            SiteTemplate = siteTemplate;
            Value = value;
        }
    }
}
