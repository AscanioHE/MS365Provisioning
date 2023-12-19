using Microsoft.Extensions.Configuration;
using MS365Provisioning.SharePoint.Model;
using MS365Provisioning.SharePoint.Services;
using MS365Provisioning.SharePoint.Settings;

namespace MS365Provisioning.SharePoint.Test
{
    public class SharePointServiceTest : ISharePointSettingsService
    {
        private readonly ISharePointService _sharePointService;
        private readonly IConfigurationRoot _config;
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
        public string? GetSetting(string key)
        {
            return _config[key];
        }
        [Fact]
        public void GetSiteSettings_ExpectDto()
        { //Act
            List<ListsSettingsDto> listsSettingsDtos = _sharePointService.GetListsSettings();
            //Assert
            Assert.NotEmpty(listsSettingsDtos);
            Assert.IsType<List<SiteSettingsDto>>(listsSettingsDtos);
        }
}