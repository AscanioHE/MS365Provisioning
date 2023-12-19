using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using MS365Provisioning.Common.Settings;
using MS365Provisioning.SharePoint.Model;
using MS365Provisioning.SharePoint.Services;
using MS365Provisioning.SharePoint.Settings;
using Xunit.Abstractions;

namespace MS365Provisioning.SharePoint.Tests
{
    public class SharePointServiceTests : IMS365ProvisioningSettings, ISharePointSettingsService
    {
        private readonly ISharePointService _sharePointService;
        private readonly IConfigurationRoot _config;

        public SharePointServiceTests(ITestOutputHelper output)
        {
            SharePointSettings sharePointSettings = new SharePointSettings();
            _config = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("dev.settings.json")
                .Build();
            ILogger logger = output.BuildLogger();
            _sharePointService = new SharePointService(this, logger,sharePointSettings.SiteUrl);
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
        [Fact]
        public void Try_ListViews_Expect_DTO()
        {
            //Act
            List<ListViewDto> listViewDtos = _sharePointService.LoadListViews();
            //Assert
            Assert.IsType<List<ListViewDto>>(listViewDtos);
        }
    }
}