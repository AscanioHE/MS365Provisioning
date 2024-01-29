using System;
using System.Collections.Generic;
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
        public uint Language { get; set; }
        public object RegionalSettings { get; set; }
        public bool QuickLaunchEnabled { get; set; }
        public bool TreeVieuwEnabled { get; set; }

        public Dictionary<string, string> Navigation { get; set; }

        public SiteSettingsDto(string title, Dictionary<string,uint> siteTemplate, string description, string logo, 
            bool siteDesignApplied, string privacySetting, bool assosiatedToHub, uint language, 
            object regionalSettings,bool quickLaunchEnabled, bool treeViewEnabled, 
            Dictionary<string, string> navigation)
        {
            Title = title;
            SiteTemplate = siteTemplate;
            Description = description;
            Logo = logo;
            PrivacySetting = privacySetting;
            SiteDesignApplied = siteDesignApplied;
            AssosiatedToHub = assosiatedToHub;
            Language = language;
            RegionalSettings = regionalSettings;
            QuickLaunchEnabled = quickLaunchEnabled;
            TreeVieuwEnabled = treeViewEnabled;
            Navigation = navigation;
        }
    }
    
}
