using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using MS365Provisioning.SharePoint.Model;
using MS365Provisioning.SharePoint.Services;
using MS365Provisioning.SharePoint.Settings;
using Xunit.Abstractions;

namespace MS365Provisioning.SharePoint.Tests
{
    public class SharePointServiceTests : ISharePointSettingsService
    {
        private readonly ISharePointService _sharePointService;
        private readonly IConfigurationRoot _config;

        public SharePointServiceTests(ITestOutputHelper output, ISharePointService sharePointService)
        {
            _sharePointService = sharePointService;
            SharePointSettings sharePointSettings = new ();
            _config = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("dev.settings.json")
                .Build();
            ILogger? logger = output.BuildLogger();
            string? siteUrl = sharePointSettings.SiteUrl;
            if (siteUrl != null) _sharePointService = new SharePointService(this, logger, siteUrl);
        }

        public string? GetSetting(string key)
        {
            return _config[key];
        }

        public SharePointSettings GetSharePointSettings()
        {
            return new SharePointSettings
            {
                ClientId = GetSetting("SharePoint:ClientId"),
                TenantId = GetSetting("SharePoint:TenantId"),
                ThumbPrint = GetSetting("SharePoint:ThumbPrint"),
                SiteUrl = GetSetting("SharePoint:SiteUrl")
            };
        }
        [Fact]

        public void Try_SiteSettings_Expect_DTO()
        {
            //Act
            List<SiteSettingsDto> siteSettingsDtos = _sharePointService.LoadSiteSettings();
            //Assert
            Assert.NotEmpty(siteSettingsDtos);
            Assert.IsType<List<SiteSettingsDto>>(siteSettingsDtos);
        }

        [Fact]
        public void Try_ListSettings_Expect_DTO()
        {
            //Act
            List<ListsSettingsDto> listsSettingsDtos = _sharePointService.LoadListsSettings();
            //Assert
            Assert.IsType<List<ListsSettingsDto>>(listsSettingsDtos);
        }
    }
}