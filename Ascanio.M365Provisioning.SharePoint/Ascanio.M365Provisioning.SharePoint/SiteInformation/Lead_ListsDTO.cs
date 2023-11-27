using Microsoft.Graph.Models;

namespace Ascanio.M365Provisioning.SharePoint.SiteInformation
{
    public class Lead_ListsDTO
    {
        public string Name { get; set; } = string.Empty;
        public string Url { get; set; } = string.Empty;
        public string ListType { get; set; } = string.Empty;
        public string ContentType { get; set; } = string.Empty;
        public bool ShowOnQuickLaunch { get; set; }
        public string QuickLaunchHeader { get; set; } = string.Empty;
        public bool AllowFolderCreation { get; set; }
        public string EnterpriseKeywords { get; set; } = string.Empty;
        public bool BreakRoleInheritance { get; set; }
        public List<Permission> Permissions { get; set; } = new List<Permission>();
    }
    
    public class Permission
    {
        public string Role { get; set; } = string.Empty;
        public bool FullControl { get; set; }
        public bool Edit {  get; set; }
        public bool Contribute { get; set; }
        public bool Delete { get; set; }
        public bool DeleteAll { get; set; }
        public bool EditAll { get; set; }
        public bool Read { get; set; }
        public bool ReadAll { get; set; }
        public bool Write { get; set; }
        public bool WriteAll { get; set; }

    }

}
