using System.Diagnostics;
using System.Security.Cryptography.X509Certificates;
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

        public SharePointService(ISharePointSettingsService sharePointSettingsService, ILogger logger, string siteUrl)
        {
            _sharePointSettingsService = sharePointSettingsService;
            _logger = logger;
            _clientContext = GetClientContext(siteUrl);
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
            ListCollection lists = _clientContext.Web.Lists;
            _clientContext.Load(lists,lc=> lc.Include(
                l=> l.Hidden)
            );
            try
            {
                _clientContext.ExecuteQuery();
                foreach (List list in lists)
                {
                    
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError($"Error fetching the ClientContext Lists: {ex.Message}");
            }

            throw new NotImplementedException();
        }

        public List<ListsSettingsDto> LoadListsSettings()
        {
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
                Debug.WriteLine($"Error creating a Certificate : {ex}");
                throw;
            }
        }
    }
}
