using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace M365Provisioning.SharePoint
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

    public class ListsSettingsDto
    {
        public string Title { get; set; } 
        public string Url { get; set; }
        public string ListType { get; set; }
        public List<string> ContentTypes { get; set; } 
        public bool ShowOnQuickLaunch { get; set; }
        public bool AllowFolderCreation { get; set; }
        public Guid EnterpriseKeywords { get; set; }
        public bool BreakRoleInheritance { get; set; }
        public Dictionary<string, string> Permissions { get; set; } 
        public List<string> QuickLauncHeaders { get; set; }

        public ListsSettingsDto(string title, string url, string listType, List<string> contentTypes,
            bool showOnQuickLaunch, List<string> quickLauncHeaders, bool allowFolderCreation, Guid enterpriseKeywords,
            bool breakRoleInheritance, Dictionary<string, string> permissions)
        {
            Title = title;
            Url = url;
            ListType = listType;
            ContentTypes = contentTypes;
            ShowOnQuickLaunch = showOnQuickLaunch;
            QuickLauncHeaders = quickLauncHeaders;
            AllowFolderCreation = allowFolderCreation;
            EnterpriseKeywords = enterpriseKeywords;
            BreakRoleInheritance = breakRoleInheritance;
            Permissions = permissions;
        }
    }
}
