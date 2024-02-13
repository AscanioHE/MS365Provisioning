using Microsoft.Extensions.Logging;
using Microsoft.IdentityModel.Tokens;
using Microsoft.SharePoint.Client;
using MS365Provisioning.Common;
using MS365Provisioning.SharePoint.Model;
using MS365Provisioning.SharePoint.Settings;
using System.Collections;
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
using User = Microsoft.SharePoint.Client.User;
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

        public ISharePointSettingsService SharePointSettingsService => _sharePointSettingsService;

        public SharePointService(ISharePointSettingsService sharePointSettingsService,
                                 ILogger logger)
        {

            sharePointSettings = new SharePointSettings();
            _sharePointSettingsService = sharePointSettingsService!;
            sharePointSettings = _sharePointSettingsService.GetSharePointSettings();
            fileSettings = _sharePointSettingsService.GetFileSettings();
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
                _logger.LogInformation($"Error fetching lists from clientcontext : {ex.Message}, StackTrace: {ex.StackTrace}");
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
                _logger?.LogError(message: $"Error fetching the Webtemplates : {ex.Message}, StackTrace: {{ex.StackTrace}}\"");
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
                                    $"Unable to create the List Data Transfer Object : {ex.Message}, StackTrace: {{ex.StackTrace}}\"");
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger?.LogInformation($"Error Fetching list properties : {ex.Message}, StackTrace: {{ex.StackTrace}}\"");
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError($"Error fetching the ClientContext Lists: {ex.Message}, StackTrace: {ex.StackTrace}");
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
                        _logger?.LogInformation($"Error fetching ClientContext: {ex}, StackTrace: {ex.StackTrace}");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.LogInformation($"Error fetching List QuickLaunchHeader : {ex.Message}, StackTrace: {ex.StackTrace} ");
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
                _logger?.LogInformation($"Error fetching Enterprise Keywords value: {ex.Message}, StackTrace: {ex.StackTrace}");
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
                _logger?.LogInformation(message: $"Error fetching permissions : {ex.Message}, StackTrace: {ex.StackTrace}");
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
                _logger?.LogInformation($"Error Fetching Lists : {ex.Message}, StackTrace: {ex.StackTrace}");
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
                        _logger?.LogInformation($"Error fetching ListView : {ex.Message}, StackTrace: {ex.StackTrace}");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.LogInformation($"Error fetching List views : {ex.Message}, StackTrace: {ex.StackTrace}");
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
                    _logger?.LogInformation($"Error fetching Site Column settings : {ex.Message}, StackTrace: {ex.StackTrace}");
                }
            }
            catch (Exception ex)
            {
                _logger?.LogInformation($"Error fetching ContextClient :  {ex.Message}, StackTrace: {ex.StackTrace}");
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
                _logger?.LogInformation($"Error fetching Content Types : {ex.Message}, StackTrace: {ex.StackTrace}");
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
                                    _logger?.LogInformation($"Error fetching SubFolders : {ex.Message}, StackTrace: {ex.StackTrace}");
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
                            _logger?.LogInformation($"Error fetching list Fields : {ex.Message}, StackTrace: {ex.StackTrace}");
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.LogInformation($"Error fetching context Lists : {ex.Message}, StackTrace: {ex.StackTrace}");
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
            SitePermissionsDto sitePermissionsDto = new();
            List<PermissionLevelDto> customPermissionLevelDtos = new();
            List<PermissionLevelDto> defaultPermissionDtos = new();
            List<string> personalPermissions = new();
            List<GroupDto> associatedGroupsDtos = new();
            List<GroupDto> groupDtos = new();
            List<UsersDto> usersDtos = new();
            List<UsersDto> siteOwnerMembers = new();
            List<string> availablePermissionLevels = new();

            Ctx.Load(Web,
                w => w.SiteGroups,
                w => w.AssociatedMemberGroup,
                w => w.AssociatedOwnerGroup,
                w => w.AssociatedVisitorGroup,
                w => w.RoleAssignments,
                w => w.RoleDefinitions,
                w => w.HasUniqueRoleAssignments);
            try
            {
                Ctx.ExecuteQuery();
                _logger?.LogWarning($"Fetching Web: Successful");
                GroupCollection siteGroups = Web.SiteGroups;
                Ctx.Load(siteGroups);
                try
                {
                    Ctx.ExecuteQuery();
                    _logger?.LogWarning($"Fetching Web.SiteGroups: Successful, Total Groups: {siteGroups.Count}");
                    foreach (Group siteGroup in siteGroups)
                    {
                        _logger?.LogWarning($"GroupName: {siteGroup.Title}");
                    }
                    foreach (Group siteGroup in siteGroups)
                    {
                        Ctx.Load(siteGroup, sg => sg.Id, sg => sg.Title);
                        try
                        {
                            Ctx.ExecuteQuery();
                            int id = siteGroup.Id;
                            _logger?.LogWarning($"GroupName: {siteGroup.Title}");
                            Group? associatedMemberGroup = Web.AssociatedMemberGroup;
                            Group? associatedOwnerGroup = Web.AssociatedOwnerGroup;
                            Group? associatedVisitorGroup = Web.AssociatedVisitorGroup;
                            Ctx.ExecuteQuery();
                            bool isAssociatedGroup = id == associatedMemberGroup.Id ||
                                                     id == associatedOwnerGroup.Id ||
                                                     id == associatedVisitorGroup.Id;
                            _logger?.LogWarning($"{siteGroup.Title} Assosciated group: {isAssociatedGroup}");
                            if (isAssociatedGroup)
                            {
                                // Check if the group already exists in the list before adding it
                                bool memberGroupExists = associatedGroupsDtos.Any(g => g.Title == associatedMemberGroup.Title);
                                bool ownerGroupExists = associatedGroupsDtos.Any(g => g.Title == associatedOwnerGroup.Title);
                                bool visitorGroupExists = associatedGroupsDtos.Any(g => g.Title == associatedVisitorGroup.Title);

                                // Add each group to the list if it doesn't already exist
                                if (!memberGroupExists)
                                {
                                    associatedGroupsDtos.Add(ConvertToGroupDto(associatedMemberGroup));
                                }
                                if (!ownerGroupExists)
                                {
                                    associatedGroupsDtos.Add(ConvertToGroupDto(associatedOwnerGroup));
                                }
                                if (!visitorGroupExists)
                                {
                                    associatedGroupsDtos.Add(ConvertToGroupDto(associatedVisitorGroup));
                                }
                            }
                            try
                            {
                                availablePermissionLevels = Web.RoleDefinitions.Select(rd => rd.Name).Distinct().ToList();
                                _logger?.LogWarning($"Fetching Available permissionlevels: Successful, Total Permissionlevels: {availablePermissionLevels.Count}");

                            }
                            catch (Exception ex)
                            {
                                _logger?.LogWarning($"Error fetching Available permissionlevels: {ex.Message}, StackTrace: {ex.StackTrace}");

                            }
                            RoleDefinitionCollection roleDefinitions = Web.RoleDefinitions;
                            foreach (RoleDefinition roleDefinition in roleDefinitions)
                            {
                                _logger?.LogInformation($"Roledefinition: {roleDefinition.Name}");
                                string groupName = GetAssignedGroup(roleDefinition);
                                if (!string.IsNullOrEmpty(groupName))
                                {
                                    Group? group = Web.SiteGroups.FirstOrDefault(grp => grp.Title == groupName);
                                    if (group != null)
                                    {
                                        groupDtos.Add(ConvertToGroupDto((Group)group));
                                    }
                                }
                                PermissionLevelDto permissionLevelDto = new()
                                {
                                    Name = roleDefinition.Name,
                                    SelectedPersonalPermissions = personalPermissions,
                                    GroupName = groupName,
                                    AssignedPermissionLevel = roleDefinition.Name,
                                    AccessRequestSettings = GetAccessRequestSettings(),
                                    SelectedListPermissions = GetSelectedListPermissions(roleDefinition)
                                };
                                if (IsDefaultPermission(roleDefinition.BasePermissions))
                                {
                                    if (!defaultPermissionDtos.Any(dto => dto.Name == permissionLevelDto.Name && dto.GroupName == permissionLevelDto.GroupName))
                                    {
                                        defaultPermissionDtos.Add(permissionLevelDto);
                                    }
                                }
                                else
                                {
                                    if (!customPermissionLevelDtos.Any(dto => dto.Name == permissionLevelDto.Name && dto.GroupName == permissionLevelDto.GroupName))
                                    {
                                        customPermissionLevelDtos.Add(permissionLevelDto);
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger?.LogWarning($"Error fetching SiteGroup {siteGroup.Title}: {ex.Message}, StackTrace: {ex.StackTrace}");
                        }
                    }

                }
                catch (Exception ex)
                {
                    _logger?.LogWarning($"Error fetching Web.SiteGroups: {ex.Message}, StackTrace: {ex.StackTrace}");
                }
                Ctx.Load(Web, w => w.SiteUsers);
                try
                {
                    Ctx.ExecuteQuery();
                    _logger?.LogWarning($"Fetching SiteUsers successful: {Web.SiteUsers.Count}");
                    List<User> siteCollectionAdmins = new();
                    foreach (User user in Web.SiteUsers)
                    {
                        Ctx.Load(user, u => u.Title, u => u.IsSiteAdmin, u => u.Email);
                        try
                        {
                            Ctx.ExecuteQuery();
                            _logger?.LogWarning($"Fetching User successful: {user.Title}");
                            if (user.IsSiteAdmin)
                            {
                                siteOwnerMembers.Add(new UsersDto
                                (
                                    user.Title,
                                    user.Email
                                ));
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger?.LogWarning($"Error fetching User: {ex.Message}, StackTrace: {ex.StackTrace}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger?.LogWarning($"Error fetching SiteUsers:  {ex.Message}, StackTrace: {ex.StackTrace}");
                }
            }
            catch (Exception ex)
            {
                _logger?.LogWarning($"Error fetching Web:  {ex.Message}, StackTrace: {ex.StackTrace}");
            }
            finally
            {
                Ctx.Dispose();
            }
            sitePermissionsDto.IsInheritedSecurity = Web.HasUniqueRoleAssignments;
            sitePermissionsDto.DefaultPermissionLevels = defaultPermissionDtos.Distinct().ToList();
            sitePermissionsDto.CustomPermissionLevels = customPermissionLevelDtos.Distinct().ToList();
            sitePermissionsDto.AvailablePermissionLevels = availablePermissionLevels;
            sitePermissionsDto.AssociatedGroups = associatedGroupsDtos.Distinct().ToList();
            sitePermissionsDto.SiteCollectionAdministrators = siteOwnerMembers;
            FileName = fileSettings!.SitePermissionsFilePath!;
            _logger?.LogInformation($"Export path to Sitesettings: {FileName}");
            DtoFile = sitePermissionsDto;
            ExportServices();
            return sitePermissionsDto;
        }


        GroupDto ConvertToGroupDto(Group group)
        {
            List<UsersDto> usersDtos = new();
            GroupDto groupDto = new();
            Ctx.Load(group,
                    g => g.Title,
                    g => g.Description,
                    g => g.LoginName,
                    g => g.Owner,
                    g => g.Users);
            Ctx.Load(group.Owner, go => go.Title);
            try
            {
                Ctx.ExecuteQuery();
                _logger?.LogWarning($"Fetching Group properties for {group.Title} Successful");
                foreach (var user in group.Users)
                {
                    Ctx.Load(user,
                                u => u.UserPrincipalName,
                                u => u.Title,
                                u => u.Email);
                    try
                    {
                        Ctx.ExecuteQuery();
                        _logger?.LogWarning($"Fetching User properties for {user.UserPrincipalName} Successful");
                        usersDtos.Add(new UsersDto(
                            user.Title,
                            user.Email
                        ));
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogWarning($"Error fetching User properties:  {ex.Message}, StackTrace: {ex.StackTrace}");
                    }
                }
                groupDto = new                
                (
                    group.Title,
                    group.Description,
                    group.LoginName,
                    group.Owner.Title,
                    usersDtos
                );
            }
            catch (Exception ex)
            {
                _logger?.LogWarning($"Error fetching Group properties:  {ex.Message}, StackTrace: {ex.StackTrace}");
            }
            return groupDto;
        }
        private string GetAccessRequestSettings()
        {
            string status = string.Empty;
            Ctx.Load(Web, w => w.RequestAccessEmail);
            try
            {
                Ctx.ExecuteQuery();
                bool accessRequestEnabled = Web.RequestAccessEmail !=null;
                if (accessRequestEnabled)
                {
                    status = "Enabled";
                }
                else
                {
                    status = "Disabled";
                }
            }
            catch (Exception ex)
            {
                _logger?.LogInformation($"Error loading Request Access Email: {ex.Message}, StackTrace: {ex.StackTrace}");
                return string.Empty;
            }
            return status;
        }

        private string GetAssignedGroup(RoleDefinition roleDefinition)
        {
            string groupName = string.Empty;
            switch (roleDefinition.Name)
            {
                case "Full Control":
                    groupName = $"{Ctx.Web.Title} Owners";
                    break;
                case "Design":
                    groupName = $"{Ctx.Web.Title} Design";
                    break;
                case "Edit":
                    groupName = $"{Ctx.Web.Title} Members";
                    break;
                case "Read":
                    groupName = $"{Ctx.Web.Title} Visitors";
                    break;
                case "Contribute":
                    groupName = $"{Ctx.Web.Title} Contributors";
                    break;
                case "Limited Access":
                    groupName = $"{Ctx.Web.Title} Limited Access";
                    break;
                default:
                    break;
            }
            // Controleer of de groep al bestaat met de berekende groepsnaam
            Group? targetGroup = Web.SiteGroups.FirstOrDefault(g => g.Title == groupName);
            Ctx.Load(targetGroup);
            Ctx.ExecuteQuery();
            if (targetGroup != null)
            {
                _logger?.LogWarning($"TargeGroup found: {groupName}");
                return groupName;
            }
            else
            {
                return string.Empty; // Groep niet gevonden
            }
        }
        private bool IsListPermission(PermissionKind permission)
        {
            // Check if the permission is related to lists
            return permission switch
            {
                PermissionKind.ManageLists => true,
                PermissionKind.EditListItems => true,
                _ => false,
            };
        }
        private bool IsDefaultPermission(BasePermissions basePermissions)
        {
            return basePermissions.Has(PermissionKind.AddListItems) &&
                   basePermissions.Has(PermissionKind.EditListItems) &&
                   basePermissions.Has(PermissionKind.DeleteListItems) &&
                   basePermissions.Has(PermissionKind.OpenItems) &&
                   basePermissions.Has(PermissionKind.ViewVersions) &&
                   basePermissions.Has(PermissionKind.DeleteVersions) &&
                   basePermissions.Has(PermissionKind.CancelCheckout) &&
                   basePermissions.Has(PermissionKind.ManagePersonalViews);
        }
        private List<string> GetSelectedListPermissions(RoleDefinition roleDefinition)
        {
            List<string> selectedListPermissions = new();
            Ctx.Load(roleDefinition, rd => rd.BasePermissions);
            Ctx.ExecuteQuery();
            _logger?.LogInformation($"Selected List Permissions:");
            if (roleDefinition.BasePermissions != null)
            {
                foreach (var permission in Enum.GetValues(typeof(PermissionKind)))
                {
                    bool isListPermission = IsListPermission((PermissionKind)permission);
                    if (roleDefinition.BasePermissions.Has((PermissionKind)permission) &&
                        isListPermission)
                    {
                        selectedListPermissions.Add(permission.ToString()!);
                        _logger?.LogInformation($"  {permission}");
                    }
                }
            }
            return selectedListPermissions;
        }
 
        //private string GetRoleDefinitionForGroup(RoleType roleType)
        //{
        //    Ctx.Load(Web.RoleAssignments,
        //     wra => wra.Include(ra => ra.RoleDefinitionBindings));
        //    Ctx.ExecuteQuery();

        //    var roleAssignment = Web.RoleAssignments.FirstOrDefault(ra =>
        //        ra.RoleDefinitionBindings.Any(rdb => rdb.RoleTypeKind == roleType));

        //    if (roleAssignment != null)
        //    {
        //        var roleDefinition = roleAssignment.RoleDefinitionBindings
        //            .FirstOrDefault(rdb => rdb.RoleTypeKind == roleType);

        //        if (roleDefinition != null)
        //        {
        //            return roleDefinition.Name;
        //        }
        //    }

        //    return string.Empty;
        //}
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
                _logger?.LogInformation($"Error fetching ContentTypes: {ex.Message}, StackTrace: {ex.StackTrace}");
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
                            _logger?.LogInformation($"Error fetching WebParts: {ex.Message}, StackTrace: {ex.StackTrace}");
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.LogInformation($"Error fetching Pages: {ex.Message}, StackTrace: {ex.StackTrace}");
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
