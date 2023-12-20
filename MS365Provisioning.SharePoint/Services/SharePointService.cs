using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using MS365Provisioning.SharePoint.Model;
using MS365Provisioning.SharePoint.Settings;
using Newtonsoft.Json;
using System.Collections;
using System.Diagnostics;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using ILogger = Microsoft.Extensions.Logging.ILogger;
using List = Microsoft.SharePoint.Client.List;

namespace MS365Provisioning.SharePoint.Services
{
    public class SharePointService : ISharePointService
    {
        private readonly ISharePointSettingsService _sharePointSettingsService;
        private readonly ILogger? _logger;
        private readonly ClientContext _clientContext;
        private readonly ListCollection? _lists;
        public SharePointService(ISharePointSettingsService? sharePointSettingsService, ILogger? logger, string siteUrl)
        {
            _sharePointSettingsService = sharePointSettingsService;
            _logger = logger;
            _clientContext = GetClientContext(siteUrl)!;
            _lists = _clientContext.Web.Lists;
        }
        /*______________________________________________________________________________________________________________
         Create ClientContext
        ________________________________________________________________________________________________________________*/
        private ClientContext? GetClientContext(string siteUrl)
        {
            string message = $"{nameof(GetClientContext)} for site {siteUrl}...";
            _logger?.LogInformation(message: message);
            if (_sharePointSettingsService != null)
            {
                SharePointSettings? sharePointSettings = _sharePointSettingsService.GetSharePointSettings();

                ClientContext? ctx;
                using (X509Certificate2? certificate = GetCertificateByThumbprint(sharePointSettings.ThumbPrint))
                {
                    ctx = null;
                    try
                    {
                        PnP.Framework.AuthenticationManager authManager = new(sharePointSettings.ClientId, certificate,
                            sharePointSettings.TenantId);
                        ctx = authManager.GetContext(sharePointSettings.SiteUrl);
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogError(message: $"Error fetching the ClientContext : {ex.Message}");
                        return new ClientContext("");
                    }
                }
                return ctx;
            }
            else
            {
                return new ClientContext("");
            }
        }
        /*______________________________________________________________________________________________________________
         Config SharePoint settings
        ________________________________________________________________________________________________________________*/
        private X509Certificate2 GetCertificateByThumbprint(string? thumbprint)
        {
            X509Certificate2 x509Certificate = new();
            try
            {
                using X509Store store = new(StoreName.My, StoreLocation.CurrentUser);
                store.Open(OpenFlags.ReadOnly);
                X509Certificate2Collection certificates = store.Certificates.Find(X509FindType.FindByThumbprint, thumbprint, false);
                if (certificates.Count > 0)
                {
                    _logger?.LogInformation("Authenticated and connected to SharePoint!");
                    x509Certificate = certificates[0];
                }
                else
                {
                    throw new InvalidOperationException($"Certificate with thumbprint {thumbprint} not found!");
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
            return siteSettingsDto;

        }

        /*______________________________________________________________________________________________________________
         Fetch Lists Settings
        ________________________________________________________________________________________________________________*/
        public List<ListsSettingsDto> LoadListsSettings()
        {
            List<ListsSettingsDto> listsSettingsDto = new();
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
            List<ListViewDto> listsSettingsDto = new();
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
                        listsSettingsDto = GetListViews(list);
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
            return listsSettingsDto;
        }
        private List<ListViewDto> GetListViews(List list)
        {
            List<ListViewDto> listviewDto = new();
            ViewCollection listViews = list.Views;
            _clientContext.Load(list,
                l=> l.Title);
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
                    catch(Exception ex)
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
                    Debug.WriteLine($"Error fetching Site Column settings : {ex.Message}");
                    throw;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error fetching ContextClient :  {ex.Message}");
                throw;
            }
            finally
            {
                _clientContext.Dispose();
            }
            return siteColumnsDtos;
        }

        public List<ContentTypesDto> LoadContentTypes()
        {
            List<ContentTypesDto> contentTypesDto = new List<ContentTypesDto>();
            try
            {
                _clientContext.Load(_lists);
                try
                {
                    _clientContext.ExecuteQuery();
                    foreach(List list in _lists)
                    {
                        if (!list.Hidden)
                        {
                            ContentTypeCollection contentTypes = list.ContentTypes;
                            _clientContext.Load(
                                contentTypes, cts=>cts.Include(
                                    ct=>ct.Name,
                                    ct=>ct.Parent,
                                    //ToDo: check required (field.Required?) 
                                    ct=>ct.Fields.Include(
                                        f=> f.InternalName)));
                            _clientContext.ExecuteQuery();
                            List<string> contentTypeFields = new ();
                            if (list.ContentTypes.Count == 0)
                            {
                                return contentTypesDto;
                            }
                            foreach (ContentType contentType in contentTypes)
                            {
                                //ToDo: check if all fields must be added
                                foreach(Field field in contentType.Fields)
                                {
                                    string fieldName = field.InternalName;
                                    contentTypeFields.Add(fieldName);
                                }
                                string contentTypeName = contentType.Name;
                                string contentTypeParent = contentType.Parent.Name;
                                
                                contentTypesDto.Add(new ContentTypesDto(
                                    contentTypeName, contentTypeParent, contentTypeFields, true));
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger?.LogInformation($"Error fetching Content Types : {ex.Message}");
                }
            }
            catch(Exception ex)
            {
                _logger?.LogInformation($"Error fetching Lists types : {ex.Message}");
            }
            finally
            {
                _clientContext.Dispose(); 
            }
            return contentTypesDto;
        }

        public List<FolderStructureDto> GetFolderStructures()
        {
            List<FolderStructureDto> folderStructureDtos = new List<FolderStructureDto>();
            _clientContext.Load(_lists);
            try
            {
                _clientContext.ExecuteQuery();

                foreach(List list in _lists)
                {
                    _clientContext.Load(
                        list,
                        l=>l.Title,
                        l=> l.Fields);
                    try
                    {
                        _clientContext.ExecuteQuery();
                        List<string> subFields = new List<string>();
                        foreach (Field field in  list.Fields)
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
            catch (Exception ex)
            {
                _logger?.LogInformation($"Error fetching context Lists : {ex.Message}");
            }
            finally
            {
                _clientContext.Dispose();
            }
            return folderStructureDtos;
        }
        private List<string> GetListContentTypes(List list)
        {
            List<string> contentTypes = new();
            try
            {
                _clientContext.Load(list.ContentTypes);
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
    }
}
