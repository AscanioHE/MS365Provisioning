using Microsoft.Graph.Models;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MS365Provisioning.SharePoint.Model
{
    public class SitePermissionsDto
    {
        public bool? IsInheritedSecurity { get; set; }
        public List<string>? AvailablePermissionLevels { get; set; }
        public List<string>? DefaultPermissionLevels { get; set; }
        public List<CustomPermissionLevelDto>? CustomPermissionLevels { get; set; }
        public List<GroupDto>? AssociatedGroups { get; set; }
        public List<string>? SiteCollectionAdministrators { get; set; }

        public SitePermissionsDto()
        {
            AvailablePermissionLevels = new List<string>();
            DefaultPermissionLevels = new List<string>();
            CustomPermissionLevels = new List<CustomPermissionLevelDto>();
            AssociatedGroups = new List<GroupDto>();
            SiteCollectionAdministrators = new List<string>();
        }
    }

    public class CustomPermissionLevelDto
    {
        public string? Name { get; set; }
        public List<string>? SelectedListPermissions { get; set; }
        public List<string>? SelectedPersonalPermissions { get; set; }
        public string? GroupName { get; set; }
        public List<string>? Members { get; set; }
        public string? AssignedPermissionLevel { get; set; }
        public bool? AccessRequestSettings { get; set; }
    }

    public class GroupDto
    {
        public string Title { get; set; }
        public string Description { get; set; }
        public int Id { get; set; }
        public bool IsHiddenInUI { get; set; }
        public string LoginName { get; set; }
        public bool AllowMembersEditMembership { get; set; }
        public bool OnlyAllowMembersViewMembership { get; set; }
        public Principal Owner { get; set; }
        public PrincipalType PrincipalType { get; set; }
        public string RequestToJoinLeaveEmailSetting { get; set; }
        public List<string> Users { get; set; }
    }
}
