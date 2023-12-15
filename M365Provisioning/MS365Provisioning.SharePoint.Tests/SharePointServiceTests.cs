using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using MS365Provisioning.Common.Settings;
using MS365Provisioning.SharePoint.Model;
using MS365Provisioning.SharePoint.Services;
using MS365Provisioning.SharePoint.Settings;
using Xunit.Abstractions;

namespace MS365Provisioning.SharePoint.Tests
{
    public class SharePointServiceTests : IMS365ProvisioningSettings, ISharePointSettingsService
    {
        //private readonly ILogger _logger;
        private readonly ISharePointService _sharePointService;
        private readonly IConfigurationRoot _config;

        public SharePointServiceTests(ITestOutputHelper output)
        {
            _config = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("dev.settings.json")
                .Build();

            ILogger logger = output.BuildLogger();
            _sharePointService = new SharePointService(this, logger, "<siteUrl>");
        }

        [Fact]
        public void Test1()
        {

        }

        public string? GetSetting(string key)
        {
            return _config[key];
        }

        public SharePointSettings GetSharePointSettings()
        {
            return new SharePointSettings
            {
                ClientId = GetSetting(""),
                TenantId = GetSetting(""),
                ThumbPrint = GetSetting(""),
                SiteUrl = GetSetting("")
            };
        }
    }
}