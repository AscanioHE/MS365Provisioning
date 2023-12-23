using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using Microsoft.SharePoint.Client.WebParts;
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
        private ClientContext Context { get; set; }
        private readonly ListCollection _lists;
        private readonly SharePointSettings sharePointSettings;
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
            SiteUrl = sharePointSettings.SiteUrl!;
            ThumbPrint = sharePointSettings.ThumbPrint!;
            Context = GetClientContext(SiteUrl);
            _logger = logger;
            DtoFile = new object();
            FileName = string.Empty;
            _lists = Context.Web.Lists;
            _lists = Context!.Web.Lists;
            try
            {
                Context.Load(_lists);
                Context.ExecuteQuery();
            }
            catch (Exception ex)
            {
                _logger.LogInformation($"Error fetching lists from clientcontext : {ex.Message}");
            }
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
            List<SiteSettingsDto> siteSettingsDto = new();
            if (sharePointSettings.SiteSettingsFilePath != null)
                FileName = sharePointSettings.SiteSettingsFilePath;
            try
            {
                WebTemplateCollection webTemplateCollection = Context.Web.GetAvailableWebTemplates(1033, true);
                Context.Load(webTemplateCollection);
                Context.ExecuteQuery();
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
                Context.Dispose();
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
            Context.Load(_lists, lc => lc.Include(
                l => l.Hidden)
                      );
            try
            {
                Context.ExecuteQuery();
                if (_lists == null || _lists.Count <= 0) return listsSettingsDto;
                foreach (List list in _lists)
                {
                    if (!list.Hidden)
                    {
                        Context.Load(
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
                            Context.ExecuteQuery();
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
                Context.Dispose();
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
                Context.Load(Context.Web.Navigation.QuickLaunch);
                Context.ExecuteQuery();
                foreach (NavigationNode navigationNode in Context.Web.Navigation.QuickLaunch)
                {
                    Context.Load
                    (
                        navigationNode,
                        n => n.Children
                    );
                    try
                    {
                        Context.ExecuteQuery();
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
                Field enterpriseKeywords = Context.Web.Fields.GetByInternalNameOrTitle("EnterpriseKeywords");
                if (enterpriseKeywords != null)
                {
                    Context.Load(enterpriseKeywords);
                    Context.ExecuteQuery();
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
            IEnumerable roles = Context.LoadQuery(queryForList);
            Dictionary<string, string> permissionDetails = new();
            try
            {
                Context.ExecuteQuery();

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
                Context.Dispose();
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
            Context.Load(_lists, lc => lc.Include(
                l => l.Hidden)
            );
            try
            {
                Context.ExecuteQuery();
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
                Context.Dispose();
            }
            DtoFile = listsViewDto;
            ExportServices();
            return listsViewDto;
        }
        private List<ListViewDto> GetListViews(List list)
        {
            List<ListViewDto> listviewDto = new();
            Microsoft.SharePoint.Client.ViewCollection listViews = list.Views;
            Context.Load(list,
                l => l.Title);
            Context.Load(listViews);
            try
            {
                Context.ExecuteQuery();
                foreach (View listView in listViews)
                {
                    Context.Load(listView);
                    Context.Load(
                        listView,
                            lv => lv.ViewFields,
                            lv => lv.Title,
                            lv => lv.DefaultView,
                            lv => lv.RowLimit,
                            lv => lv.Scope);
                    try
                    {
                        Context.ExecuteQuery();
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
                Context.Dispose();
            }
            return listviewDto;
        }
        /*______________________________________________________________________________________________________________
         Fetch Lists SiteColumns
        ________________________________________________________________________________________________________________*/
        public List<SiteColumnsDto> LoadSiteColumns()
        {
            List<SiteColumnsDto> siteColumnsDtos = new();
            if (sharePointSettings.SiteColumnsFilePath != null)
                FileName = sharePointSettings.SiteColumnsFilePath;
            try
            {
                FieldCollection siteColumns = Context.Web.Fields;
                Context.Load(siteColumns,
                             scc => scc.Include(
                                    sc => sc.Hidden,
                                    sc => sc.InternalName,
                                    sc => sc.SchemaXml,
                                    sc => sc.DefaultValue));
                try
                {
                    Context.ExecuteQuery();
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
                Context.Dispose();
            }
            DtoFile = siteColumnsDtos;
            ExportServices();
            return siteColumnsDtos;
        }

        public List<ContentTypesDto> LoadContentTypes()
        {
            List<ContentTypesDto> contentTypesDto = new();
            if (sharePointSettings.ContentTypesFilePath != null)
                FileName = sharePointSettings.ContentTypesFilePath;
            try
            {
                foreach (List list in _lists)
                {
                    if (!list.Hidden)
                    {
                        ContentTypeCollection contentTypes = list.ContentTypes;
                        Context.Load(
                            contentTypes, cts => cts.Include(
                                ct => ct.Name,
                                ct => ct.Parent,
                                ct => ct.Fields.Include(
                                    f => f.InternalName)));
                        Context.ExecuteQuery();
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
                Context.Dispose();
            }
            DtoFile = contentTypesDto;
            ExportServices();
            return contentTypesDto;
        }

        public List<FolderStructureDto> GetFolderStructures()
        {
            List<FolderStructureDto> folderStructureDtos = new();
            if (sharePointSettings.FolderStructureFilePath != null)
                FileName = sharePointSettings.FolderStructureFilePath;
            try
            {
                foreach (List list in _lists)
                {
                    if (!list.Hidden)
                    {
                        Context.Load(
                            list,
                            l => l.Title,
                            l => l.Fields);
                        try
                        {
                            Context.ExecuteQuery();
                            List<string> subFields = new();
                            foreach (Field field in list.Fields)
                            {
                                Context.Load(field,
                                    f => f.Title);
                                try
                                {
                                    Context.ExecuteQuery();
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
                Context.Dispose();
            }
            DtoFile = folderStructureDtos;
            ExportServices();
            return folderStructureDtos;
        }

        public List<SitePermissionsDto> LoadSitePermissions()
        {
            List<SitePermissionsDto> sitePermissionsDtos = new();
            if (sharePointSettings.SitePermissionsFilePath != null)
                try
                {
                    Context.Load(Context.Web,
                        w => w.Title,
                        w => w.SiteGroups.Include(
                            item => item.Users,
                            item => item.PrincipalType,
                            item => item.LoginName,
                            item => item.Title));
                    Context.ExecuteQuery();
                    string webTitle = Context.Web.Title;

                    foreach (Group siteGroup in Context.Web.SiteGroups.Where(group => group.Title.Contains(webTitle)))
                    {
                        List<string> userNames = new();
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
            FileName = sharePointSettings!.SitePermissionsFilePath;
            DtoFile = sitePermissionsDtos;
            ExportServices();
            return sitePermissionsDtos;
        }
        private List<string> GetListContentTypes(List list)
        {
            List<string> contentTypes = new();
            try
            {
                Context!.Load(list.ContentTypes);
                Context.ExecuteQuery();
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
            foreach (List list in _lists)
            {
                if (!list.Hidden)
                {
                    var absoluteUrl = "https://23m2yz.sharepoint.com/sites/M365ProvisioningLve";
                    var relativeUrl = absoluteUrl.Replace("https://23m2yz.sharepoint.com/", "");

                    var testSiteUrl = "https://23m2yz.sharepoint.com/sites/TestSite1";
                    var testWebParts = Context.Web.GetWebParts(relativeUrl).ToList();

                    foreach (var v in testWebParts)
                    {
                        var s = v;
                        // Voer hier de gewenste bewerkingen uit op de naam van het webonderdeel (s)
                    }
                }
                    
            }
            DtoFile = webPartPagesDtos;
            ExportServices();
            return webPartPagesDtos;
        }
        public void ExportServices()
        {
            ExportServices exportServices = new()
            {
                DtoFile = DtoFile
            };
            exportServices.FileName = exportServices.ConvertToJsonString();
            exportServices.WriteJsonStringToFile();
        }
    }
}
