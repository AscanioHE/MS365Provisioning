using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using MS365Provisioning.Common;
using MS365Provisioning.SharePoint.Model;
using MS365Provisioning.SharePoint.Settings;
using System.Collections;
using System.Security.Cryptography.X509Certificates;
using System.Text;

namespace MS365Provisioning.SharePoint.Services
{
    public class SharePointService : ISharePointService
    {
        private readonly ISharePointSettingsService _sharePointSettingsService;
        private readonly ILogger _logger;
        private ClientContext _clientContext { get; set; }
        private readonly ListCollection _lists;
        private readonly SharePointSettings sharePointSettings;
        private object DtoFile;
        private string FileName { get; set; }
        private string ThumbPrint { get; set; }
        private string SiteUrl { get; set; }

        public ISharePointSettingsService SharePointSettingsService => _sharePointSettingsService;

        public SharePointService(ISharePointSettingsService sharePointSettingsService,
                                 ILogger logger,
                                 string siteUrl)
        {

            sharePointSettings= new SharePointSettings();
            SiteUrl = siteUrl;
            SiteUrl = sharePointSettings.SiteUrl!;
            _clientContext = new ClientContext(sharePointSettings.SiteUrl);
            _sharePointSettingsService = sharePointSettingsService!;
            sharePointSettings = _sharePointSettingsService.GetSharePointSettings();
            _logger = logger;
            DtoFile = new object();
            FileName = string.Empty;
            _lists = _clientContext.Web.Lists;
            ThumbPrint = sharePointSettings.ThumbPrint!;
            _clientContext = GetClientContext(siteUrl);
            _lists = _clientContext!.Web.Lists;
            _clientContext.Load(_lists);
            _clientContext.ExecuteQuery();
        }
        /*______________________________________________________________________________________________________________
         Create ClientContext
        ________________________________________________________________________________________________________________*/
        private ClientContext GetClientContext(string siteUrl)
        {
            string message = $"{nameof(GetClientContext)} for site {siteUrl}...";
            _logger?.LogInformation(message: message);
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
            X509Certificate2 x509Certificate = new();
            try
            {
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
            }
            catch (Exception ex)
            {
                _logger?.LogInformation($"Error creating a Certificate : {ex}");
            }
            return x509Certificate;
        }
        /*______________________________________________________________________________________________________________
         Fetch SiteSettings
        ________________________________________________________________________________________________________________*/

        public List<SiteSettingsDto> LoadSiteSettings()
        {
            List<SiteSettingsDto> siteSettingsDto = new();
            if (sharePointSettings.SiteSettingsFilePath != null)
                FileName = sharePointSettings.SiteSettingsFilePath;
            try
            {
                WebTemplateCollection webTemplateCollection = _clientContext.Web.GetAvailableWebTemplates(1033, true);
                _clientContext.Load(webTemplateCollection);
                _clientContext.ExecuteQuery();
                foreach (WebTemplate webTemplate in webTemplateCollection)
                {
                    siteSettingsDto.Add(new SiteSettingsDto
                    {
                        SiteTemplate = webTemplate.Name,
                        Value = webTemplate.Lcid
                    });
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError(message: $"Error fetching the Webtemplates : {ex.Message}");
            }
            finally
            {
                _clientContext.Dispose();
            }
            DtoFile = siteSettingsDto;
            ExportServices();
            return siteSettingsDto;
        }

        /*______________________________________________________________________________________________________________
         Fetch Lists Settings
        ________________________________________________________________________________________________________________*/
        public List<ListsSettingsDto> LoadListsSettings()
        {
            List<ListsSettingsDto> listsSettingsDto = new();
            if (sharePointSettings.ListsFilePath != null)
                FileName = sharePointSettings.ListsFilePath;
            bool breakRoleAssignment = false;
            _clientContext.Load(_lists, lc => lc.Include(
                l => l.Hidden)
                      );
            try
            {
                _clientContext.ExecuteQuery();
                if (_lists == null || _lists.Count <= 0) return listsSettingsDto;
                foreach (List list in _lists)
                {
                    if (!list.Hidden)
                    {
                        _clientContext.Load(
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
                            _clientContext.ExecuteQuery();
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
                _clientContext.Dispose();
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
                _clientContext.Load(_clientContext.Web.Navigation.QuickLaunch);
                _clientContext.ExecuteQuery();
                foreach (NavigationNode navigationNode in _clientContext.Web.Navigation.QuickLaunch)
                {
                    _clientContext.Load
                    (
                        navigationNode,
                        n => n.Children
                    );
                    try
                    {
                        _clientContext.ExecuteQuery();
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
                Field enterpriseKeywords = _clientContext.Web.Fields.GetByInternalNameOrTitle("EnterpriseKeywords");
                if (enterpriseKeywords != null)
                {
                    _clientContext.Load(enterpriseKeywords);
                    _clientContext.ExecuteQuery();
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
            IEnumerable roles = _clientContext.LoadQuery(queryForList);
            Dictionary<string, string> permissionDetails = new();
            try
            {
                _clientContext.ExecuteQuery();

                foreach (RoleAssignment ra in roles)
                {
                    RoleDefinitionBindingCollection rdc = ra.RoleDefinitionBindings;
                    StringBuilder permissionBuilder = new();
                    foreach (RoleDefinition rd in rdc)
                    {
                        permissionBuilder.Append(rd.Name + ", ");
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
                _clientContext.Dispose();
            }
            return permissionDetails;
        }
        /*______________________________________________________________________________________________________________
         Fetch Lists List Views
        ________________________________________________________________________________________________________________*/
        public List<ListViewDto> LoadListViews()
        {
            List<ListViewDto> listsViewDto = new();
            if (sharePointSettings.ListViewsFilePath != null)
                FileName = sharePointSettings.ListViewsFilePath;
            _clientContext.Load(_lists, lc => lc.Include(
                l => l.Hidden)
            );
            try
            {
                _clientContext.ExecuteQuery();
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
                _clientContext.Dispose();
            }
            DtoFile = listsViewDto;
            ExportServices();
            return listsViewDto;
        }
        private List<ListViewDto> GetListViews(List list)
        {
            List<ListViewDto> listviewDto = new();
            Microsoft.SharePoint.Client.ViewCollection listViews = list.Views;
            _clientContext.Load(list,
                l => l.Title);
            _clientContext.Load(listViews);
            try
            {
                _clientContext.ExecuteQuery();
                foreach (View listView in listViews)
                {
                    _clientContext.Load(listView);
                    _clientContext.Load(
                        listView,
                            lv => lv.ViewFields,
                            lv => lv.Title,
                            lv => lv.DefaultView,
                            lv => lv.RowLimit,
                            lv => lv.Scope);
                    try
                    {
                        _clientContext.ExecuteQuery();
                        List<string> viewFields = new List<string>();
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
                _clientContext.Dispose();
            }
            return listviewDto;
        }
        /*______________________________________________________________________________________________________________
         Fetch Lists SiteColumns
        ________________________________________________________________________________________________________________*/
        public List<SiteColumnsDto> LoadSiteColumns()
        {
            List<SiteColumnsDto> siteColumnsDtos = new List<SiteColumnsDto>();
            if (sharePointSettings.SiteColumnsFilePath != null)
                FileName = sharePointSettings.SiteColumnsFilePath;
            try
            {
                FieldCollection siteColumns = _clientContext.Web.Fields;
                _clientContext.Load(siteColumns,
                             scc => scc.Include(
                                    sc => sc.Hidden,
                                    sc => sc.InternalName,
                                    sc => sc.SchemaXml,
                                    sc => sc.DefaultValue));
                try
                {
                    _clientContext.ExecuteQuery();
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
                _clientContext.Dispose();
            }
            DtoFile = siteColumnsDtos;
            ExportServices();
            return siteColumnsDtos;
        }

        public List<ContentTypesDto> LoadContentTypes()
        {
            List<ContentTypesDto> contentTypesDto = new List<ContentTypesDto>();
            if (sharePointSettings.ContentTypesFilePath != null)
                FileName = sharePointSettings.ContentTypesFilePath;
            try
            {
                _clientContext.Load(_lists);
                try
                {
                    _clientContext.ExecuteQuery();
                    foreach (List list in _lists)
                    {
                        if (!list.Hidden)
                        {
                            ContentTypeCollection contentTypes = list.ContentTypes;
                            _clientContext.Load(
                                contentTypes, cts => cts.Include(
                                    ct => ct.Name,
                                    ct => ct.Parent,
                                    ct => ct.Fields.Include(
                                        f => f.InternalName)));
                            _clientContext.ExecuteQuery();
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
            }
            catch (Exception ex)
            {
                _logger?.LogInformation($"Error fetching Lists types : {ex.Message}");
            }
            finally
            {
                _clientContext.Dispose();
            }
            DtoFile = contentTypesDto;
            ExportServices();
            return contentTypesDto;
        }

        public List<FolderStructureDto> GetFolderStructures()
        {
            List<FolderStructureDto> folderStructureDtos = new List<FolderStructureDto>();
            if (sharePointSettings.FolderStructureFilePath != null)
                FileName = sharePointSettings.FolderStructureFilePath;
            _clientContext.Load(_lists);
            try
            {
                _clientContext.ExecuteQuery();

                foreach (List list in _lists)
                {
                    if (!list.Hidden)
                    {
                        _clientContext.Load(
                            list,
                            l => l.Title,
                            l => l.Fields);
                        try
                        {
                            _clientContext.ExecuteQuery();
                            List<string> subFields = new List<string>();
                            foreach (Field field in list.Fields)
                            {
                                _clientContext.Load(field,
                                    f => f.Title);
                                try
                                {
                                    _clientContext.ExecuteQuery();
                                    subFields.Add(field.Title);
                                }
                                catch (Exception ex)
                                {
                                    _logger?.LogInformation($"Error fetching SubFolders : {ex.ToString()}");
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
                _clientContext.Dispose();
            }
            DtoFile = folderStructureDtos;
            ExportServices();
            return folderStructureDtos;
        }

        public List<SitePermissionsDto> LoadSitePermissions()
        {
            List<SitePermissionsDto> sitePermissionsDtos = new();
            if (sharePointSettings.SitePermissionsFilePath != null)
                FileName = sharePointSettings!.SitePermissionsFilePath;
            try
            {
                _clientContext.Load(_clientContext.Web,
                    w => w.Title,
                    w => w.SiteGroups.Include(
                        item => item.Users,
                        item => item.PrincipalType,
                        item => item.LoginName,
                        item => item.Title));
                _clientContext.ExecuteQuery();
                string webTitle = _clientContext.Web.Title;

                foreach (Group siteGroup in _clientContext.Web.SiteGroups.Where(group => group.Title.Contains(webTitle)))
                {
                    List<string> userNames = new List<string>();
                    foreach (User user in siteGroup.Users)
                    {
                        if (!user.IsHiddenInUI && user.PrincipalType == PrincipalType.User)
                        {
                            userNames.Add(user.UserPrincipalName);
                        }
                    }
                    sitePermissionsDtos.Add(new SitePermissionsDto
                        (webTitle, siteGroup.Title, userNames));
                }
            }
            catch (Exception ex)
            {
                _logger?.LogInformation($"Error fetching Site Permissions : {ex.Message}");
            }
            DtoFile = sitePermissionsDtos;
            ExportServices();
            return sitePermissionsDtos;
        }
        private List<string> GetListContentTypes(List list)
        {
            List<string> contentTypes = new();
            try
            {
                _clientContext!.Load(list.ContentTypes);
                _clientContext.ExecuteQuery();
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
        public void ExportServices()
        {
            ExportServices exportServices = new();
            exportServices.DtoFile = DtoFile;
            exportServices.FileName = exportServices.ConvertToJsonString();
            exportServices.WriteJsonStringToFile();
        }
    }
}
