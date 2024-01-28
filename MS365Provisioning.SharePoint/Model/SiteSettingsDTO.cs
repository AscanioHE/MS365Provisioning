namespace MS365Provisioning.SharePoint.Model
{
    public class SiteSettingsDto
    {
        public Dictionary<string, uint> SiteTemplate { get; set; }
        public uint Value { get; set; }
        public string Title { get; set; }
        public string Description { get; set; }
        public string Logo { get; set; }
        public bool SiteDesignApplied { get; set; }
        public string PrivacySetting { get; set; }  
        public bool AssosiatedToHub { get; set; }

        public SiteSettingsDto(Dictionary<string,uint> siteTemplate, uint value, string title, string description, string logo, bool siteDesignApplied, string privacySetting, bool assosiatedToHub)
        {
            SiteTemplate = siteTemplate;
            Value = value;
            Title = title;
            Description = description;
            Logo = logo;
            SiteDesignApplied = siteDesignApplied;
            PrivacySetting = privacySetting;
            AssosiatedToHub = assosiatedToHub;
        }
    }

}
