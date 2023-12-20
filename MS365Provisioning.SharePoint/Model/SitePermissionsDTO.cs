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
        public SitePermissionsDto(string groupName,string permissionLevel,List<string> userNames) 
        {
            GroupName = groupName;
            PermissionLevel = permissionLevel;
            UserNames = userNames;
        }
    }
}
