using Ascanio.M365Provisioning.SharePoint.SiteInformation;
using Microsoft.Extensions.Logging;
using Xunit.Abstractions;

namespace M365Provisioning.SharePoint.Tests
{
    public class SiteSettingsTests
    {
        private readonly ILogger _logger;
        private readonly ISiteSettingsService _siteSettingsService;
        public SiteSettingsTests(ITestOutputHelper output)
        {
            _logger = output.BuildLogger();
            _siteSettingsService = new SiteSettings();
        }

        [Fact]
        public void Try_GetSiteSettings_Expect_DTO()
        {
            //Arrange
            ISiteSettingsService siteSettingsService = new SiteSettings();

            //Act
            List<SiteSettingsDTO> siteSettings = siteSettingsService.Load();

            //Assert
            Assert.IsType<SiteSettingsDTO>(siteSettings);
            Assert.IsType<List<SiteSettingsDTO>>(siteSettings);
            Assert.True(siteSettings.Any());
        }

        [Fact]
        public void Try_ConvertToJson_Expect_jsonString()
        {
            //Arrange
            List<SiteSettingsDTO> siteSettings = _siteSettingsService.Load();
            string json = _siteSettingsService.ConvertToJson(siteSettings);

            //Assert
            Assert.IsType<string>(json);
        }
    }
}