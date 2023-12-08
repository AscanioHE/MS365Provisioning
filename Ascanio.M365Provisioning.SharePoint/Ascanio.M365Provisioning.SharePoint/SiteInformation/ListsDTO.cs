using Microsoft.Graph.Models;

namespace Ascanio.M365Provisioning.SharePoint.SiteInformation
{
    public class ListsDTO
    {
        public string Title { get; set; } = string.Empty;
        public string Url { get; set; } = string.Empty;
        public string ListType { get; set; } = string.Empty;
        public List<string> ContentTypes { get; set; } = new List<string>();
        public bool ShowOnQuickLaunch { get; set; }
        public bool AllowFolderCreation { get; set; }
        public Guid EnterpriseKeywords { get; set; } 
        public bool BreakRoleInheritance { get; set; }
        public Dictionary<string, string> Permissions { get; set; } = new Dictionary<string, string>();

        public ListsDTO(string title, string url, string listType, List<string> contentTypes,
                             bool showOnQuickLaunch, bool allowFolderCreation, Guid enterpriseKeywords,
                             bool breakRoleInheritance, Dictionary<string, string> permissions)
        {
            Title = title;
            Url = url;
            ListType = listType;
            ContentTypes = contentTypes;
            ShowOnQuickLaunch = showOnQuickLaunch;
            AllowFolderCreation = allowFolderCreation;
            EnterpriseKeywords = enterpriseKeywords;
            BreakRoleInheritance = breakRoleInheritance;
            Permissions = permissions;
        }
        public ListsDTO() { }
    }
}
