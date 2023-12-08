using Ascanio.M365Provisioning.SharePoint.Services;
using Ascanio.M365Provisioning.SharePoint.SiteInformation;
using Microsoft.SharePoint.Client;

namespace Ascanio.M365Provisioning.SharePoint
{
    public class Program
    {
        private static ISiteSettingsService _siteSettingsService;

        static void Main()
        {

            SharePointService _spService = new();
            _siteSettingsService = new SiteSettings();

            List<SiteSettingsDTO> siteSettings = _siteSettingsService.Load();
            if (siteSettings.Any())
            {
                _siteSettingsService.WriteToJsonFile(siteSettings, _spService.SiteSettingsFilePath);
            }

            Run();
        }
        static void Run()
        {
            _ = new SharePointFunction();
        }
    }
}