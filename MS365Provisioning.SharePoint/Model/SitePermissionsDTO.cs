using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MS365Provisioning.SharePoint.Model
{
    public class SitePermissionsDto
    {
        public string GroupName { get; set; }
        public string PermissionLevel { get; set; }
        public List<string> UserNames { get; set; }
        public bool IsInherited { get; set; }
        public string SecurityType { get; set; }
        public bool IsDefault { get; set; }
        public bool IsCustom { get; set; }
        public bool IsAssociatedWithSite { get; set; }
        public List<string> ListPermissions { get; set; }
        public List<string> SitePermissions { get; set; }
        public List<string> PersonalPermissions { get; set; }
        public string AccessRequestSettings { get; set; }
        public List<string> SiteCollectionAdministrators { get; set; }
        public SitePermissionsDto(string groupName,string permissionLevel,List<string> userNames,
                                   bool isInherited,string securityType,bool isDefault,bool isCustom,
                                   bool isAssociatedWithSite,List<string> listPermissions,List<string> sitePermissions,
                                   List<string> personalPermissions,string accessRequestSettings,List<string> siteCollectionAdministrators) 
        {
            GroupName = groupName;
            PermissionLevel = permissionLevel;
            UserNames = userNames;
            IsInherited = isInherited;
            SecurityType = securityType;
            IsDefault = isDefault;
            IsCustom = isCustom;
            IsAssociatedWithSite = isAssociatedWithSite;
            ListPermissions = listPermissions;
            SitePermissions = sitePermissions;
            AccessRequestSettings = accessRequestSettings;
            SiteCollectionAdministrators = siteCollectionAdministrators;
        }
    }
}
