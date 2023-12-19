namespace MS365Provisioning.SharePoint.Model
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
