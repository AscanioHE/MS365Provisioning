using M365Provisioning.SharePoint.Services;
using M365Provisioning.SharePoint.DTO;
using M365Provisioning.SharePoint;
using Microsoft.Extensions.Logging;
using PnP.Framework.Provisioning.Model;
using Xunit.Abstractions;

namespace M365Provisioning.test
{
    public class SharePointSettingsTests
    {
        public readonly ILogger _logger;
        public readonly ISharePointServices _sharePointService;
        public readonly WriteData2Json _writeData2Json;

        public SharePointSettingsTests(ITestOutputHelper output)
        {
            _logger = output.BuildLogger();
            _sharePointService = new SharePointServices();
            _writeData2Json = new WriteData2Json();
        }

        [Fact]
        public void Try_GetSiteSettings_Expect_DTO()
        {
            //Arrange
            ISharePointServices sharePointService = new SharePointServices();

            //Act
            List<SiteSettingsDto> siteSettings = sharePointService.GetSiteSettings();

            //Assert
            Assert.IsType<SiteSettingsDto>(siteSettings);
        }

        [Fact]
        public void Try_ConvertToJson_Expect_jsonString()
        {
            //Arrange
            List<SiteSettingsDto> siteSettings = _sharePointService.GetSiteSettings();
            string json = _writeData2Json.ConvertToJson(siteSettings);

            //Assert
            Assert.IsType<string>(json);
        }
    }
}