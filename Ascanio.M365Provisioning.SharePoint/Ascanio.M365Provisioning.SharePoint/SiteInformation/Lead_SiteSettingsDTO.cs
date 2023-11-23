namespace Ascanio.M365Provisioning.SharePoint.SiteInformation
{
    public class Lead_SiteSettingsDTO
    {
        public string SiteTemplate { get; set; } = string.Empty;
        public string Value { get; set; } = string.Empty;
        public Lead_SiteSettingsDTO() { }
        public Lead_SiteSettingsDTO(string siteTemplate, string value)
        {
            SiteTemplate = siteTemplate;
            Value = value;
        }
    }
}
