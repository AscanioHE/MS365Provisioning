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
        public List<UsersDto>? SiteCollectionAdministrators { get; set; }

        public SitePermissionsDto() { }
        public SitePermissionsDto(List<string> availablePermissionLevels,List<PermissionLevelDto> defaultPermissionLevelDtos,List<PermissionLevelDto> customPermissionLevel,
            List<GroupDto> associatedGroups,List<UsersDto> usersDtos)
        {
            AvailablePermissionLevels = availablePermissionLevels;
            DefaultPermissionLevels = defaultPermissionLevelDtos; 
            CustomPermissionLevels = customPermissionLevel; 
            AssociatedGroups = associatedGroups;
            SiteCollectionAdministrators = usersDtos;
        }
    }

    public class PermissionLevelDto
    {
        public string? Name { get; set; }
        public List<string>? SelectedPersonalPermissions { get; set; }
        public string? GroupName { get; set; }
        public string? AssignedPermissionLevel { get; set; }
        public string? AccessRequestSettings { get; set; }
        public List<string>? SelectedListPermissions { get; set; }
        public PermissionLevelDto() { }
        public PermissionLevelDto(string name, List<string> selectedPersonalPermissions, string groupName,string assignedPermissionLevel,
            string accessRequestSettings,List<string> selectedListPermissions)
        {
            Name = name;
            SelectedPersonalPermissions = selectedPersonalPermissions;
            GroupName = groupName;
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
        public string Owner { get; set; }
        public List<UsersDto> Members { get; set; }
        // Standaardconstructor toegevoegd
        public GroupDto()
        {
            Title = "";
            Description = "";
            LoginName = "";
            Owner = "";
            Members = new List<UsersDto>();
        }
        public GroupDto(string title, string description, string loginName,string owner,List<UsersDto> members) 
        { 
            Title = title;
            Description = description;
            LoginName = loginName;
            Owner = owner;
            Members = members;
        }
    }
    public class UsersDto
    {
        public string Email { get; set; }
        public string Title { get; set; }
        public UsersDto(string email, string title)
        {
            Email = email;
            Title = title;
        }
    }
}
