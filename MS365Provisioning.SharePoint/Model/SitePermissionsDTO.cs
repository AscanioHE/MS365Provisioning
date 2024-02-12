using Microsoft.Graph.Models;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using Newtonsoft.Json;
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
        public List<PermissionLevelDto>? DefaultPermissionLevels { get; set; }
        public List<PermissionLevelDto>? CustomPermissionLevels { get; set; }
        public List<GroupDto>? AssociatedGroups { get; set; }
        public List<Users>? SiteCollectionAdministrators { get; set; }

        public SitePermissionsDto()
        {
            AssociatedGroups = new List<GroupDto>();
            AvailablePermissionLevels = new List<string>();
            SiteCollectionAdministrators = new List<Users>();
            DefaultPermissionLevels = new List<PermissionLevelDto>();
            CustomPermissionLevels = new List<PermissionLevelDto>();
        }
    }

    public class PermissionLevelDto
    {
        public string? Name { get; set; }
        public List<string>? SelectedPersonalPermissions { get; set; }
        public string? GroupName { get; set; }
        public List<Users>? Members { get; set; }
        public string? AssignedPermissionLevel { get; set; }
        public string? AccessRequestSettings { get; set; }
        public List<string>? SelectedListPermissions { get; set; }
        public PermissionLevelDto(string name, List<string> selectedPersonalPermissions, string groupName, List<Users> members, string assignedPermissionLevel,
            string accessRequestSettings,List<string> selectedListPermissions)
        {
            Name = name;
            SelectedPersonalPermissions = selectedPersonalPermissions;
            GroupName = groupName;
            Members = members;
            AssignedPermissionLevel = assignedPermissionLevel;
            AccessRequestSettings = accessRequestSettings;
            SelectedListPermissions = selectedListPermissions;
        }
    }

    public class GroupDto
    {
        public string Title { get; set; }
        public string Description { get; set; }
        public string LoginName { get; set; }
        public Principal Owner { get; set; }
        public List<Users> Members { get; set; }
        public GroupDto(string title, string description, string loginName, Principal owner, List<Users> members) 
        { 
            Title = title;
            Description = description;
            LoginName = loginName;
            Owner = owner;
            Members = members;
        }
    }
    public class Users
    {
        public string UserPrincipalName { get; set; }
        public string Email { get; set; }
        public string Title { get; set; }
        public bool IsSiteAdmin { get; set; }
        public Users(string userPrincipalName, string email, string title, bool isSiteAdmin)
        {
            UserPrincipalName = userPrincipalName;
            Email = email;
            Title = title;
            IsSiteAdmin = isSiteAdmin;
        }
    }
}
