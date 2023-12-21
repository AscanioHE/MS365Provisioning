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

        public SharePointServiceTests(ITestOutputHelper output)
        {

            _config = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("dev.settings.json")
                .Build();
            SharePointSettings sharePointSettings = new();
            ILogger? logger = output.BuildLogger();
            string siteUrl = sharePointSettings.SiteUrl!;
            _sharePointService = new SharePointService(this,logger,siteUrl);
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
                SiteUrl = GetSetting("SharePoint:SiteUrl"),
                FolderStructureFilePath = GetSetting("SharePoint:FolderStructureFilePath"),
                ListsFilePath = GetSetting("SharePoint:ListsFilePath"),
                ListViewsFilePath = GetSetting("SharePoint:ListViewsFilePath"),
                SiteColumnsFilePath = GetSetting("SharePoint:SiteColumnsFilePath"),
                SiteSettingsFilePath = GetSetting("SharePoint:SiteSettingsFilePath"),
                SitePermissionsFilePath = GetSetting("SharePoint:SitePermissionsFilePath"),
                WebPartsFilePath = GetSetting("SharePoint:WebPartsFilePath"),
                ContentTypesFilePath = GetSetting("SharePoint:ContentTypesFilePath")
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
            Assert.NotEmpty(listsSettingsDtos);
            Assert.IsType<List<ListsSettingsDto>>(listsSettingsDtos);
        }
        [Fact]
        public void Try_ListViews_Expect_DTO()
        {
            //Act
            List<ListViewDto> listViewDtos = _sharePointService.LoadListViews();
            //Assert
            Assert.NotEmpty(listViewDtos);
            Assert.IsType<List<ListViewDto>>(listViewDtos);
        }
        [Fact]
        public void Try_LoadContentTypes_Expect_DTO()
        {
            //Act
            List<ContentTypesDto> contentTypesDto = _sharePointService.LoadContentTypes();
            //Assert
            Assert.NotEmpty(contentTypesDto);
            Assert.IsType<List<ContentTypesDto>>(contentTypesDto);
        }
        [Fact]
        public void Try_SiteColumns_Expect_DTO()
        {
            //Act
            List<SiteColumnsDto> siteColumnsDto = _sharePointService.LoadSiteColumns();
            //Assert
            Assert.NotEmpty(siteColumnsDto);
            Assert.IsType<List<SiteColumnsDto>>(siteColumnsDto);
        }
        [Fact]
        public void Try_FolderStructure_Expect_DTO()
        {
            //Act
            List<FolderStructureDto> folderStructureDto = _sharePointService.GetFolderStructures();
            //Assert
            Assert.NotEmpty(folderStructureDto);
            Assert.IsType<List<FolderStructureDto>>(folderStructureDto);
        }
        [Fact]
        public void Try_LoadSitePermissions_Expect_DTO()
        {
            //Act
            List<SitePermissionsDto> sitePermissions = _sharePointService.LoadSitePermissions();
            //Assert
            Assert.NotEmpty(sitePermissions);
            Assert.IsType<List<SitePermissionsDto>>(sitePermissions);
        }
    }
}