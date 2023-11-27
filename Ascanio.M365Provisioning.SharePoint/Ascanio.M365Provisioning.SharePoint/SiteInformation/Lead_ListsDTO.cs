using Microsoft.Graph.Models;
using Microsoft.SharePoint.Client.Sharing;

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
        public Lead_ListsDTO(string name, string url, string listType, string contentType,
                             bool showOnQuickLaunch,bool allowFolderCreation, string enterpriseKeyWords,
                             bool breakRoleInheritiance) 
        {
            Name = name;
            Url = url;
            ListType = listType;
            ContentType = contentType;
            ShowOnQuickLaunch = showOnQuickLaunch;
            AllowFolderCreation = allowFolderCreation;
            EnterpriseKeywords = enterpriseKeyWords;
            BreakRoleInheritance = breakRoleInheritiance;

        }
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
        public Permission(string role, bool fullControl, bool edit, 
                          bool editAll, bool contribute,bool delete, bool deleteAll,
                          bool read, bool readAll, bool write, bool writeAll)
        {
            Role = role;
            FullControl = fullControl;
            Edit = edit;
            EditAll = editAll;
            Contribute = contribute;
            Delete = delete;
            DeleteAll = deleteAll;
            Read = read;
            ReadAll = readAll;
            Write = write;
            WriteAll = writeAll;
        }
    }
}
