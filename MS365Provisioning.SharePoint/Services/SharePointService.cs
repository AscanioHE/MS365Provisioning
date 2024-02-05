using Microsoft.Extensions.Logging;
using Microsoft.IdentityModel.Tokens;
using Microsoft.SharePoint.Client;
using MS365Provisioning.Common;
using MS365Provisioning.SharePoint.Model;
using MS365Provisioning.SharePoint.Settings;
using PnP.Core.Model.SharePoint;
using System.Collections;
using System.Diagnostics.CodeAnalysis;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using ContentType = Microsoft.SharePoint.Client.ContentType;
using ContentTypeCollection = Microsoft.SharePoint.Client.ContentTypeCollection;
using Context = Microsoft.SharePoint.Client.ClientContext;
using Field = Microsoft.SharePoint.Client.Field;
using FieldCollection = Microsoft.SharePoint.Client.FieldCollection;
using Group = Microsoft.SharePoint.Client.Group;
using List = Microsoft.SharePoint.Client.List;
using ListItem = Microsoft.SharePoint.Client.ListItem;
using NavigationNode = Microsoft.SharePoint.Client.NavigationNode;
using PermissionKind = Microsoft.SharePoint.Client.PermissionKind;
using RoleAssignment = Microsoft.SharePoint.Client.RoleAssignment;
using RoleAssignmentCollection = Microsoft.SharePoint.Client.RoleAssignmentCollection;
using RoleDefinition = Microsoft.SharePoint.Client.RoleDefinition;
using RoleType = Microsoft.SharePoint.Client.RoleType;
using View = Microsoft.SharePoint.Client.View;
using WebPart = Microsoft.SharePoint.Client.WebParts.WebPart;

namespace MS365Provisioning.SharePoint.Services
{
    public class SharePointService : ISharePointService
    {
        private readonly ISharePointSettingsService _sharePointSettingsService;
        private readonly ILogger _logger;
        private Context Ctx { get; set; }
        private Web Web { get; set; }
        private readonly ListCollection _lists;
        private readonly SharePointSettings sharePointSettings;
        private readonly FileSettings fileSettings;
        private object DtoFile;
        private string FileName { get; set; }
        private string ThumbPrint { get; set; }
        private string SiteUrl { get; set; }

        public ISharePointSettingsService SharePointSettingsService => _sharePointSettingsService;

        public SharePointService(ISharePointSettingsService sharePointSettingsService,
                                 ILogger logger)
        {

            sharePointSettings = new SharePointSettings();
            _sharePointSettingsService = sharePointSettingsService!;
            sharePointSettings = _sharePointSettingsService.GetSharePointSettings();
            fileSettings = _sharePointSettingsService.GetFileSettings();
            SiteUrl = sharePointSettings.SiteUrl!;
            ThumbPrint = sharePointSettings.ThumbPrint!;
            Ctx = GetClientContext();
            Web = Ctx.Web;
            Ctx.Load(Web);
            Ctx.ExecuteQuery();
            _logger = logger;
            DtoFile = new object();
            FileName = string.Empty;
            _lists = Ctx.Web.Lists;
            _lists = Ctx!.Web.Lists;
            try
            {
                Ctx.Load(_lists);
                Ctx.ExecuteQuery();
            }
            catch (Exception ex)
            {
                _logger.LogInformation($"Error fetching lists from clientcontext : {ex.Message}");
            }
        }
        /*______________________________________________________________________________________________________________
         Create ClientContext
        ________________________________________________________________________________________________________________*/
        private ClientContext GetClientContext()
        {
            ClientContext ctx;
            X509Certificate2 certificate = GetCertificateByThumbprint();
            PnP.Framework.AuthenticationManager authManager = new(sharePointSettings.ClientId, certificate,
                sharePointSettings.TenantId);
            ctx = authManager.GetContext(sharePointSettings.SiteUrl);
            return ctx;
        }
        /*______________________________________________________________________________________________________________
         Config SharePoint settings
        ________________________________________________________________________________________________________________*/
        private X509Certificate2 GetCertificateByThumbprint()
        {
            X509Certificate2 x509Certificate;
            using X509Store store = new(StoreName.My, StoreLocation.CurrentUser);
            store.Open(OpenFlags.ReadOnly);
            X509Certificate2Collection certificates = store.Certificates.Find(X509FindType.FindByThumbprint, ThumbPrint, false);
            if (certificates.Count > 0)
            {
                _logger?.LogInformation("Authenticated and connected to SharePoint!");
                x509Certificate = certificates[0];
            }
            else
            {
                throw new InvalidOperationException($"Certificate with thumbprint {ThumbPrint} not found!");
            }
            return x509Certificate;
        }
        /*______________________________________________________________________________________________________________
         Fetch SiteSettings
        ________________________________________________________________________________________________________________*/

        public List<SiteSettingsDto> LoadSiteSettings()
        {
            List<SiteSettingsDto> siteSettingsDtos = new();
            Web web = Ctx.Web;
            Ctx.Load(Ctx.Web,
                w => w.Title,
                w => w.Url,
                w => w.Description,
                w => w.SiteLogoUrl,
                w => w.WebTemplate,
                w => w.Title,
                w => w.RelatedHubSiteIds,
                w => w.Language,
                w => w.RegionalSettings,
                w => w.Navigation,
                w => w.QuickLaunchEnabled,
                w => w.TreeViewEnabled,
                w => w.HeaderLayout,
                w => w.CustomMasterUrl,
                w => w.Navigation.QuickLaunch,
                w => w.Navigation.TopNavigationBar);
            try
            {
                ObjectSharingSettings objectSharingSettings = web.GetObjectSharingSettingsForSite(true);
                var sharingSettings = web.GetObjectSharingSettingsForSite;
                bool privacySettings = sharingSettings.Method.IsPublic;
                Ctx.Load(web.RegionalSettings, rs => rs.TimeZone.Id, rs => rs.DateFormat, rs => rs.LocaleId, rs => rs.TimeZone.Description);
                Ctx.ExecuteQuery();

                string title = web.Title;
                string url = web.Url;
                string description = web.Description;
                string currentWebTemplate = web.WebTemplate;
                string logo = web.SiteLogoUrl;
                bool siteDesignApplied = web.WebTemplate != "STS";
                string privacy = privacySettings ? "Public" : "Private";
                var relatedHubSiteIds = web.RelatedHubSiteIds;
                bool assosiatedToHub = !relatedHubSiteIds.IsNullOrEmpty();
                var regionalSettings =
                (
                    web.RegionalSettings.DateFormat,
                    web.RegionalSettings.TimeZone.Description,
                    web.RegionalSettings.LocaleId
                );
                uint language = web.Language;
                Dictionary<string, string> navigationItems = new();
                var navigation = web.Navigation;
                foreach (var node in navigation.QuickLaunch)
                {
                    navigationItems.Add(node.Title, node.Url);
                }
                foreach (var node in web.Navigation.TopNavigationBar)
                {
                    navigationItems.Add(node.Title, node.Url);
                }
                bool quickLaunchEnabled = web.QuickLaunchEnabled;
                bool treeViewEnabled = web.TreeViewEnabled;
                string headerLayout = web.HeaderLayout.ToString();

                Dictionary<string, uint> webTemplates = new();
                if (fileSettings.SiteSettingsFilePath != null)
                {
                    FileName = fileSettings.SiteSettingsFilePath;
                }
                WebTemplateCollection webTemplateCollection = Ctx.Web.GetAvailableWebTemplates(1033, true);
                Ctx.Load(webTemplateCollection);
                Ctx.ExecuteQuery();
                foreach (WebTemplate webTemplate in webTemplateCollection)
                {
                    if (!webTemplates.ContainsKey(webTemplate.Title))
                    {
                        webTemplates.Add(webTemplate.Title, webTemplate.Lcid);
                    }
                }

                siteSettingsDtos.Add(new SiteSettingsDto
                    (
                        title,
                        description,
                        currentWebTemplate,
                        logo,
                        siteDesignApplied,
                        privacy,
                        assosiatedToHub,
                        language,
                        regionalSettings,
                        quickLaunchEnabled,
                        treeViewEnabled,
                        navigationItems,
                        headerLayout,
                        webTemplates
                    ));

            }
            catch (Exception ex)
            {
                _logger?.LogError(message: $"Error fetching the Webtemplates : {ex.Message}");
            }
            finally
            {
                Ctx.Dispose();
            }
            DtoFile = siteSettingsDtos;
            ExportServices();
            return siteSettingsDtos;
        }

        /*______________________________________________________________________________________________________________
         Fetch Lists Settings
        ________________________________________________________________________________________________________________*/
        public List<ListsSettingsDto> LoadListsSettings()
        {
            List<ListsSettingsDto> listsSettingsDto = new();
            FileName = fileSettings.ListsFilePath!;
            bool breakRoleAssignment = false;
            Ctx.Load(_lists, lc => lc.Include(
                l => l.Hidden)
                      );
            try
            {
                Ctx.ExecuteQuery();
                if (_lists == null || _lists.Count <= 0) return listsSettingsDto;
                foreach (List list in _lists)
                {
                    if (!list.Hidden)
                    {
                        Ctx.Load(
                            list,
                            l => l.Title,
                            l => l.DefaultViewUrl,
                            l => l.BaseType,
                            l => l.ContentTypes,
                            l => l.OnQuickLaunch,
                            l => l.HasUniqueRoleAssignments,
                            l => l.EnableFolderCreation,
                            l => l.RoleAssignments,
                            l => l.Fields.Include(
                                f => f.InternalName,
                                f => f.Title));
                        try
                        {
                            Ctx.ExecuteQuery();
                            List<string> contentTypes = GetListContentTypes(list);
                            Dictionary<string, string> listPermissions = GetPermissionDetails(list);
                            Guid enterpiseKeywordsValue = GetEnterpriseKeywordsValue();
                            List<string> quickLaunchHeaders = GetQuickLaunchHeaders();
                            RoleAssignmentCollection roleAssignmentCollection = list.RoleAssignments;
                            foreach (RoleAssignment roleAssignment in roleAssignmentCollection)
                            {
                                breakRoleAssignment = roleAssignment.RoleDefinitionBindings.AreItemsAvailable;
                            }
                            try
                            {
                                listsSettingsDto.Add(new
                                (
                                    list.Title,
                                    list.DefaultViewUrl,
                                    list.BaseType.ToString(),
                                    contentTypes,
                                    list.OnQuickLaunch,
                                    quickLaunchHeaders,
                                    list.EnableFolderCreation,
                                    enterpiseKeywordsValue,
                                    breakRoleAssignment,
                                    listPermissions
                                ));
                            }
                            catch (Exception ex)
                            {
                                _logger?.LogInformation(
                                    $"Unable to create the List Data Transfer Object : {ex.Message}");
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger?.LogInformation($"Error Fetching list properties : {ex.Message}");
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError($"Error fetching the ClientContext Lists: {ex.Message}");
            }
            finally
            {
                Ctx.Dispose();
            }
            DtoFile = listsSettingsDto;
            ExportServices();
            return listsSettingsDto;
        }

        private List<string> GetQuickLaunchHeaders()
        {
            List<string> quickLaunchHeaders = new();
            try
            {
                Ctx.Load(Ctx.Web.Navigation.QuickLaunch);
                Ctx.ExecuteQuery();
                foreach (NavigationNode navigationNode in Ctx.Web.Navigation.QuickLaunch)
                {
                    Ctx.Load
                    (
                        navigationNode,
                        n => n.Children
                    );
                    try
                    {
                        Ctx.ExecuteQuery();
                        foreach (NavigationNode childNode in navigationNode.Children)
                        {
                            quickLaunchHeaders.Add(childNode.Title);
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogInformation($"Error fetching ClientContext: {ex}");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.LogInformation($"Error fetching List QuickLaunchHeader : {ex.Message}");
            }

            return quickLaunchHeaders;
        }
        private Guid GetEnterpriseKeywordsValue()
        {
            Guid enterpriseKeywordsValue = Guid.Empty;
            try
            {
                Field enterpriseKeywords = Ctx.Web.Fields.GetByInternalNameOrTitle("EnterpriseKeywords");
                if (enterpriseKeywords != null)
                {
                    Ctx.Load(enterpriseKeywords);
                    Ctx.ExecuteQuery();
                    enterpriseKeywordsValue = enterpriseKeywords.Id;
                }
            }
            catch (Exception ex)
            {
                _logger?.LogInformation($"Error fetching Enterprise Keywords value: {ex.Message}");
                enterpriseKeywordsValue = Guid.Empty;
            }
            return enterpriseKeywordsValue;
        }
        Dictionary<string, string> GetPermissionDetails(List list)
        {
            IQueryable<RoleAssignment> queryForList = list.RoleAssignments.Include(
                roleAsg => roleAsg.Member,
                roleAsg => roleAsg.RoleDefinitionBindings.Include(roleDef => roleDef.Name));
            IEnumerable roles = Ctx.LoadQuery(queryForList);
            Dictionary<string, string> permissionDetails = new();
            try
            {
                Ctx.ExecuteQuery();

                foreach (RoleAssignment ra in roles)
                {
                    RoleDefinitionBindingCollection rdc = ra.RoleDefinitionBindings;
                    StringBuilder permissionBuilder = new();
                    foreach (RoleDefinition rd in rdc)
                    {
                        permissionBuilder.Append(rd.Name);
                        _logger?.LogInformation(permissionBuilder.ToString());
                    }
                    string permission = permissionBuilder.ToString();
                    permissionBuilder.Clear();

                    permissionDetails.Add(permission, ra.Member.Title);
                }
            }
            catch (Exception ex)
            {
                _logger?.LogInformation(message: $"Error fetching permissions : {ex}");
            }
            finally
            {
                Ctx.Dispose();
            }
            return permissionDetails;
        }
        /*______________________________________________________________________________________________________________
         Fetch Lists List Views
        ________________________________________________________________________________________________________________*/
        public List<ListViewDto> LoadListViews()
        {
            List<ListViewDto> listsViewDto = new();
            if (fileSettings.ListViewsFilePath != null)
                FileName = fileSettings.ListViewsFilePath;
            Ctx.Load(_lists, lc => lc.Include(
                l => l.Hidden)
            );
            try
            {
                Ctx.ExecuteQuery();
                foreach (List list in _lists)
                {
                    if (!list.Hidden)
                    {
                        listsViewDto = GetListViews(list);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.LogInformation($"Error Fetching Lists : {ex.Message}");
            }
            finally
            {
                Ctx.Dispose();
            }
            DtoFile = listsViewDto;
            ExportServices();
            return listsViewDto;
        }
        private List<ListViewDto> GetListViews(List list)
        {
            List<ListViewDto> listviewDto = new();
            Microsoft.SharePoint.Client.ViewCollection listViews = list.Views;
            Ctx.Load(list,
                l => l.Title);
            Ctx.Load(listViews);
            try
            {
                Ctx.ExecuteQuery();
                foreach (View listView in listViews)
                {
                    Ctx.Load(listView);
                    Ctx.Load(
                        listView,
                            lv => lv.ViewFields,
                            lv => lv.Title,
                            lv => lv.DefaultView,
                            lv => lv.RowLimit,
                            lv => lv.Scope);
                    try
                    {
                        Ctx.ExecuteQuery();
                        List<string> viewFields = new();
                        foreach (string field in listView.ViewFields)
                        {
                            viewFields.Add(field);
                        }
                        listviewDto.Add(new(
                            list.Title,
                            listView.Title,
                            listView.DefaultView,
                            viewFields,
                            listView.RowLimit,
                            listView.Scope.ToString(),
                            $"{list.Title}"
                            ));
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogInformation($"Error fetching ListView : {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.LogInformation($"Error fetching Listviews : {ex.Message}");
            }
            finally
            {
                Ctx.Dispose();
            }
            return listviewDto;
        }
        /*______________________________________________________________________________________________________________
         Fetch Lists SiteColumns
        ________________________________________________________________________________________________________________*/
        public List<SiteColumnsDto> LoadSiteColumns()
        {
            List<SiteColumnsDto> siteColumnsDtos = new();
            if (fileSettings.SiteColumnsFilePath != null)
                FileName = fileSettings.SiteColumnsFilePath;
            try
            {
                FieldCollection siteColumns = Ctx.Web.Fields;
                Ctx.Load(siteColumns,
                             scc => scc.Include(
                                    sc => sc.Hidden,
                                    sc => sc.InternalName,
                                    sc => sc.SchemaXml,
                                    sc => sc.DefaultValue));
                try
                {
                    Ctx.ExecuteQuery();
                    foreach (Field siteColumn in siteColumns)
                    {
                        siteColumnsDtos.Add(new SiteColumnsDto(
                            siteColumn.InternalName, siteColumn.SchemaXml, siteColumn.DefaultValue));
                    }
                }
                catch (Exception ex)
                {
                    _logger?.LogInformation($"Error fetching Site Column settings : {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                _logger?.LogInformation($"Error fetching ContextClient :  {ex.Message}");
            }
            finally
            {
                Ctx.Dispose();
            }
            DtoFile = siteColumnsDtos;
            ExportServices();
            return siteColumnsDtos;
        }

        public List<ContentTypesDto> LoadContentTypes()
        {
            List<ContentTypesDto> contentTypesDto = new();
            if (fileSettings.ContentTypesFilePath != null)
                FileName = fileSettings.ContentTypesFilePath;
            try
            {
                foreach (List list in _lists)
                {
                    if (!list.Hidden)
                    {
                        ContentTypeCollection contentTypes = list.ContentTypes;
                        Ctx.Load(
                            contentTypes, cts => cts.Include(
                                ct => ct.Name,
                                ct => ct.Parent,
                                ct => ct.Fields.Include(
                                    f => f.InternalName)));
                        Ctx.ExecuteQuery();
                        List<string> contentTypeFields = new();
                        if (list.ContentTypes.Count == 0)
                        {
                            return contentTypesDto;
                        }
                        foreach (ContentType contentType in contentTypes)
                        {
                            contentTypeFields.AddRange(
                                    from Field field in contentType.Fields
                                    let fieldName = field.InternalName
                                    select fieldName);
                            string contentTypeName = contentType.Name;
                            string contentTypeParent = contentType.Parent.Name;

                            contentTypesDto.Add(new ContentTypesDto(
                                contentTypeName, contentTypeParent, contentTypeFields));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.LogInformation($"Error fetching Content Types : {ex.Message}");
            }
            finally
            {
                Ctx.Dispose();
            }
            DtoFile = contentTypesDto;
            ExportServices();
            return contentTypesDto;
        }

        public List<FolderStructureDto> GetFolderStructures()
        {
            List<FolderStructureDto> folderStructureDtos = new();
            if (fileSettings.FolderStructureFilePath != null)
                FileName = fileSettings.FolderStructureFilePath;
            try
            {
                foreach (List list in _lists)
                {
                    if (!list.Hidden)
                    {
                        Ctx.Load(
                            list,
                            l => l.Title,
                            l => l.Fields);
                        try
                        {
                            Ctx.ExecuteQuery();
                            List<string> subFields = new();
                            foreach (Field field in list.Fields)
                            {
                                Ctx.Load(field,
                                    f => f.Title);
                                try
                                {
                                    Ctx.ExecuteQuery();
                                    subFields.Add(field.Title);
                                }
                                catch (Exception ex)
                                {
                                    _logger?.LogInformation($"Error fetching SubFolders : {ex.Message}");
                                }
                                folderStructureDtos.Add(new(
                                    list.Title,
                                    field.Title,
                                    subFields
                                    ));
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger?.LogInformation($"Error fetching list Fields : {ex.Message}");
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.LogInformation($"Error fetching context Lists : {ex.Message}");
            }
            finally
            {
                Ctx.Dispose();
            }
            DtoFile = folderStructureDtos;
            ExportServices();
            return folderStructureDtos;
        }

        public SitePermissionsDto LoadSitePermissions()
        {
            List<SitePermissionsDto> sitePermissionsDtos = new();

            Ctx.Load(Web.AssociatedOwnerGroup);
            Ctx.Load(Web.AssociatedMemberGroup);
            Ctx.Load(Web.AssociatedVisitorGroup);
            Ctx.Load(Web.RoleAssignments);
            Ctx.Load(Web.RoleDefinitions);
            Ctx.Load(Ctx.Site.RootWeb,
                rw=>rw.HasUniqueRoleAssignments,
                rw=>rw.RequestAccessEmail
                );
            Ctx.Load(Web.SiteGroups,
                    sg => sg.Include
                    (
                        g => g.Title,
                        g => g.Description,
                        g => g.Id,
                        g => g.IsHiddenInUI,
                        g => g.LoginName,
                        g=> g.AllowMembersEditMembership,
                        g => g.OnlyAllowMembersViewMembership,
                        g => g.Owner,
                        g => g.PrincipalType,
                        g => g.RequestToJoinLeaveEmailSetting,
                        g => g.Users.Include
                        (
                            u=>u.UserPrincipalName
                            )
                        )
                    ); ;
            Ctx.ExecuteQuery();
            List<string> siteOwnerMembers = new List<string>();
            var siteOwners = Web.AssociatedOwnerGroup.Users;
            Ctx.Load(siteOwners);
            Ctx.ExecuteQuery();

            SitePermissionsDto sitePermissionsDto = new SitePermissionsDto();
            //Site Administrators
            UserCollection Owners = siteOwners;
            siteOwnerMembers.AddRange(siteOwners.Select(user => user.UserPrincipalName));
            _logger?.LogInformation($"Site Owners:");
            foreach (string member in siteOwnerMembers)
            {
                _logger?.LogInformation($"  {member}");
            }

            // Available Permission Levels
            List<string> availablePermissionLevels = Web.RoleDefinitions.Select(rd => rd.Name).ToList();
            sitePermissionsDto.AvailablePermissionLevels = availablePermissionLevels;
            _logger?.LogInformation($"Available Permission levels:");
            foreach (string availablePermissionLevel in availablePermissionLevels)
            {
                _logger?.LogInformation($"  {availablePermissionLevel}");
            }
            List<GroupDto> groupDtos = new List<GroupDto>();
            foreach (Group group in Web.SiteGroups)
            {
                GroupDto groupDto = new GroupDto
                {
                    Title = group.Title,
                    Description = group.Description,
                    Id = group.Id,
                    IsHiddenInUI = group.IsHiddenInUI,
                    LoginName = group.LoginName,
                    AllowMembersEditMembership = group.AllowMembersEditMembership,
                    OnlyAllowMembersViewMembership = group.OnlyAllowMembersViewMembership,
                    Owner = group.Owner,
                    PrincipalType = group.PrincipalType,
                    RequestToJoinLeaveEmailSetting = group.RequestToJoinLeaveEmailSetting,
                    Users = GetGroupMembers(group) // Veronderstellend dat GetGroupMembers een methode is die de gebruikers van de groep ophaalt
                };
                    _logger?.LogInformation($"{groupDto}");
                    groupDtos.Add(groupDto);

                _logger?.LogInformation($"Permissions for group: {groupDto.Title}");

                foreach (RoleType roleType in Enum.GetValues(typeof(RoleType)))
                {
                    string permissionLevel = GetRoleDefinitionForGroup(group.Title, roleType);

                    if (!string.IsNullOrEmpty(permissionLevel))
                    {
                        _logger?.LogInformation($"RoleType: {roleType}, PermissionsLevel: {permissionLevel}");
                    }
                    else
                    {
                        _logger?.LogInformation($"RoleType: {roleType}, Permission level not found for group {group.Title}");
                    }
                }
            }
            // Custom Permission Levels
            List<CustomPermissionLevelDto> customPermissionLevelDtos = new();
            foreach (RoleDefinition roleDefinition in Web.RoleDefinitions)
            {
                CustomPermissionLevelDto customPermissionLevelDto = new CustomPermissionLevelDto
                {
                    Name = roleDefinition.Name,
                    GroupName = GetAssignedGroup(roleDefinition),
                    SelectedListPermissions = GetSelectedListPermissions(roleDefinition),
                    AssignedPermissionLevel = roleDefinition.Name,
                    AccessRequestSettings = GetAccessRequestSettings()
                };
                customPermissionLevelDtos.Add(customPermissionLevelDto);
            }

            sitePermissionsDto.SiteCollectionAdministrators = siteOwnerMembers;
            sitePermissionsDto.CustomPermissionLevels = customPermissionLevelDtos;
            FileName = fileSettings!.SitePermissionsFilePath!;
            DtoFile = sitePermissionsDto;
            ExportServices();
            return sitePermissionsDto;
        }

        private bool GetAccessRequestSettings()
        {
            foreach (List list in _lists)
            {
                if (Ctx.Site.RootWeb.HasUniqueRoleAssignments)
                {
                    return !string.IsNullOrEmpty(Ctx.Site.RootWeb.RequestAccessEmail);
                }
            }
            return false;
        }

        private string GetAssignedGroup(RoleDefinition roleDefinition)
        {
            string groupName = string.Empty;
            switch(roleDefinition.Name)
            {
                case "Full Control":
                    groupName = $"{Ctx.Web.Title} Owners";
                    break;
                case "Edit":
                    groupName = $"{Ctx.Web.Title} Members";
                    break;
                case "Read":
                    groupName = $"{Ctx.Web.Title} Visitors";
                    break;

                default:
                    break;
                
            }
            GroupCollection groups = Web.SiteGroups;
            Ctx.Load(
                        groups,gc=>gc.Include
                        (
                            g=>g.LoginName,
                            g=>g.Title,
                            g=>g.PrincipalType
                        )
                    );
            Ctx.ExecuteQuery();
            foreach(Group group in Web.SiteGroups)
            {
                if(group.LoginName == groupName)
                {
                    return groupName;
                }
            }
            //string groupName = groups.LoginName;
            return string.Empty;
        }
        private bool IsListPermission(PermissionKind permission)
        {
            // Check if the permission is related to lists
            switch (permission)
            {
                case PermissionKind.AddListItems:
                case PermissionKind.EditListItems:
                case PermissionKind.DeleteListItems:
                case PermissionKind.OpenItems:
                case PermissionKind.ViewVersions:
                case PermissionKind.DeleteVersions:
                case PermissionKind.CancelCheckout:
                case PermissionKind.ManagePersonalViews:
                    return true;
                default:
                    return false;
            }
        }
        private List<string> GetSelectedListPermissions(RoleDefinition roleDefinition)
        {
            List<string> selectedListPermissions = new List<string>();
            Ctx.Load(roleDefinition, rd => rd.BasePermissions);
            Ctx.ExecuteQuery();
            _logger?.LogInformation($"Selected List Permissions:");
            if (roleDefinition.BasePermissions != null)
            {
                foreach (var permission in Enum.GetValues(typeof(PermissionKind)))
                {
                    if (roleDefinition.BasePermissions.Has((PermissionKind)permission) &&
                        IsListPermission((PermissionKind)permission))
                    {
                        selectedListPermissions.Add(permission.ToString()!);
                        _logger?.LogInformation($"  {permission.ToString()}");
                    }
                }
            }
            return selectedListPermissions;
        }
        private List<string> GetGroupMembers(Group group)
        {
            List<string> users= new();
            foreach(User user in group.Users)
            {
                users.Add(user.UserPrincipalName);
            }
            return users;
        }
        private string GetRoleDefinitionForGroup(string groupName, RoleType roleType)
        {
            Ctx.Load(Web.RoleAssignments,
             wra => wra.Include(ra => ra.RoleDefinitionBindings));
            Ctx.ExecuteQuery();

            var roleAssignment = Web.RoleAssignments.FirstOrDefault(ra =>
                ra.RoleDefinitionBindings.Any(rdb => rdb.RoleTypeKind == roleType));

            if (roleAssignment != null)
            {
                var roleDefinition = roleAssignment.RoleDefinitionBindings
                    .FirstOrDefault(rdb => rdb.RoleTypeKind == roleType);

                if (roleDefinition != null)
                {
                    return roleDefinition.Name;
                }
            }

            return string.Empty;
        }
        private List<string> GetListContentTypes(List list)
        {
            List<string> contentTypes = new();
            try
            {
                Ctx!.Load(list.ContentTypes);
                Ctx.ExecuteQuery();
                if (list.ContentTypes.Count == 0)
                {
                    return contentTypes;
                }
                foreach (ContentType contentType in list.ContentTypes)
                {
                    contentTypes.Add(contentType.Name);
                }
            }
            catch (Exception ex)
            {
                _logger?.LogInformation($"Error fetching ContentTypes: {ex.Message}");
                contentTypes.Clear();
            }
            return contentTypes;
        }
        public List<WebPartPagesDto> LoadWebParts()
        {
            List<WebPartPagesDto> webPartPagesDtos = new();
            List<WebPart> webparts = new();
            try
            {
                List pagesList = Ctx.Web.Lists.GetByTitle("Site Pages");
                CamlQuery camlQuery = new();
                Ctx.Load(pagesList);
                Ctx.ExecuteQuery();
                ListItemCollection pages = pagesList.GetItems(camlQuery);
                Ctx.Load(pages);
                Ctx.ExecuteQuery();
                foreach (ListItem item in pages)
                {
                    Ctx.Load(item, I => I.DisplayName, I => I.File);
                    Ctx.ExecuteQueryRetry();
                    if (item.DisplayName == "Home")
                    {
                        var file = item.File;
                        Ctx.Load(file);
                        Ctx.ExecuteQuery();
                        var page = Ctx.Web.LoadClientSidePage(item.DisplayName);
                        Ctx.ExecuteQuery();
                        var webParts = page.Controls;
                        if (webParts != null && webParts.Count > 0)
                        {
                            foreach (var control in webparts)
                            {
                                foreach (object o in control.Properties.FieldValues)
                                {
                                    var i = o.GetType().Name;
                                    var j = o.ToString();
                                }

                            }
                        }
                        try
                        {

                        }
                        catch (Exception ex)
                        {
                            _logger?.LogInformation($"Error fetching WebParts: {ex.Message}");
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.LogInformation($"Error fetching Pages: {ex.Message}");
            }

            DtoFile = webPartPagesDtos;
            ExportServices();
            return webPartPagesDtos;
        }
        public void ExportServices()
        {
            ExportServices exportServices = new()
            {
                DtoFile = DtoFile,
                FileName = FileName,
            };
            exportServices.ConvertToJsonString();
            exportServices.WriteJsonStringToFile();
        }
    }
}
