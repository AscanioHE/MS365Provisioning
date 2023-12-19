using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using MS365Provisioning.SharePoint.Model;
using MS365Provisioning.SharePoint.Settings;
using System.Collections;
using System.Security.Cryptography.X509Certificates;
using System.Text;
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
        public SharePointService(ClientContext clientContext)
        {
            _clientContext = clientContext;
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
            try
            {
                using X509Store store = new(StoreName.My, StoreLocation.CurrentUser);
                store.Open(OpenFlags.ReadOnly);
                var certificates = store.Certificates.Find(X509FindType.FindByThumbprint, thumbprint, false);
                if (certificates.Count > 0)
                {
                    _logger?.LogInformation("Authenticated and connected to SharePoint!");
                    return certificates[0];
                }
                else
                {
                    throw new InvalidOperationException($"Certificate with thumbprint {thumbprint} not found!");
                }
            }
            catch (Exception ex)
            {
                _logger?.LogInformation($"Error creating a Certificate : {ex}");
                return null;
            }
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
            List<ListsSettingsDto> list = new List<ListsSettingsDto>();
            return list;
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
            try
            {
                _clientContext.ExecuteQuery();

                Dictionary<string, string> permissionDetails = new();
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
                return permissionDetails;
            }
            catch (Exception ex)
            {
                _logger?.LogInformation(message: $"Error fetching permissions : {ex}");
                return new Dictionary<string, string>();
            }
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
                _logger.LogInformation($"Error Fetching Lists : {ex.Message}");
            }
            return listsSettingsDto;
        }
        private List<ListViewDto> GetListViews(List list)
        {
            List<ListViewDto> listviewDto = new();
            ViewCollection listViews = list.Views;
            _clientContext.Load(listViews);
            try
            {
                _clientContext.ExecuteQuery();
                foreach (View listView in listViews)
                {
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
            }
            catch (Exception ex)
            {
                _logger.LogInformation($"Error fetching Listviews : {ex.Message}");
            }
            return listviewDto;
        }
        /*______________________________________________________________________________________________________________
         Fetch Lists SiteColumns
        ________________________________________________________________________________________________________________*/
        public List<SiteColumnsDto> LoadSiteColumnsDtos()
        {
            List<SiteColumnsDto> list = new List<SiteColumnsDto>();
            return list;
        }
    }
}
