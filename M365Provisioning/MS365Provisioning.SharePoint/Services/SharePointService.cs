using System.Collections;
using System.Diagnostics;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using MS365Provisioning.SharePoint.Model;
using MS365Provisioning.SharePoint.Settings;

namespace MS365Provisioning.SharePoint.Services
{
    public class SharePointService : ISharePointService
    {
        private readonly ISharePointSettingsService _sharePointSettingsService;
        private readonly ILogger _logger;
        private readonly ClientContext _clientContext;
        private readonly ListCollection _lists;

        public SharePointService(ISharePointSettingsService sharePointSettingsService, ILogger logger, string siteUrl)
        {
            _sharePointSettingsService = sharePointSettingsService;
            _logger = logger;
            _clientContext = GetClientContext(siteUrl)!; 
            _lists = _clientContext.Web.Lists;

        }

        public SharePointService(ClientContext? clientContext)
        {
            _clientContext = clientContext;
        }

        private ClientContext? GetClientContext(string siteUrl)
        {
            _logger?.LogInformation($"{nameof(GetClientContext)} for site {siteUrl}...");

            SharePointSettings sharePointSettings = _sharePointSettingsService.GetSharePointSettings();

            X509Certificate2 certificate = GetCertificateByThumbprint(sharePointSettings.ThumbPrint);
            ClientContext? ctx = null;

            try
            {
                PnP.Framework.AuthenticationManager authManager = new(sharePointSettings.ClientId, certificate, sharePointSettings.TenantId);
                ctx = authManager.GetContext(sharePointSettings.SiteUrl);
            }
            catch (Exception ex)
            {
                _logger?.LogError($"Error fetching the ClientContext : {ex.Message}");
            }

            return ctx;
        }

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
                _logger?.LogError($"Error fetching the Webtemplates : {ex.Message}");
            }
            finally
            {
                _clientContext.Dispose();
            }
            return siteSettingsDto;
        }

        public List<ListsSettingsDto> LoadListsSettings()
        {
            _clientContext.Load(_lists, lc => lc.Include(
                l => l.Hidden)
            );
            try
            {
                _clientContext.ExecuteQuery();
            }
            catch (Exception ex)
            {
                _logger?.LogError($"Error fetching the ClientContext Lists: {ex.Message}");
            }
            throw new NotImplementedException();
        }

        public List<ListViewDto> LoadListViews()
        {
            throw new NotImplementedException();
        }

        public List<SiteColumnsDto> LoadSiteColumnsDtos()
        {
            throw new NotImplementedException();
        }

      
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
                throw;
            }
        }
        private List<string> GetListContentTypes(List list)
        {
            List<string> contentTypes = new();

            try
            {
                if (list.ContentTypes.Count == 0)
                {
                    return contentTypes; // No ContentTypes to return
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
        private List<string> GetQuickLaunchHeaders()
        {
            List<string> quickLaunchHeaders = new();
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
                    //_logger?.LogInformation($"Error fetching ClientContext: {ex}");
                    throw;
                }
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
                // Log the exception
                //_logger?.LogInformation($"Error fetching Enterprise Keywords value: {ex.Message}");
            }

            return enterpriseKeywordsValue;
        }
        Dictionary<string, string> GetPermissionDetails(ClientContext _clientContext, IQueryable<RoleAssignment> queryString)
        {
            {
                IEnumerable roles = _clientContext.LoadQuery(queryString);
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
                    _logger?.LogInformation($"Error fetching permissions : {ex}");
                    return new Dictionary<string, string>();
                }
            }
        }
    }
}
