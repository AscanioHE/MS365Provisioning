using M365Provisioning.SharePoint.DTO;
using M365Provisioning.SharePoint.Functions;
using M365Provisioning.SharePoint.Interfaces;
using M365Provisioning.SharePoint.Services;
using Microsoft.SharePoint.Client;

namespace M365Provisioning.Test
{
    public class SharePointTest
    {
        private ISharePointServices SharePointServices { get; set; } = new SharePointServices();
        private ISharePointFunctions SharePointFunctions { get; set; } = new SharePointFunctions();


        [Fact]
        public void Try_GetClientContext_Expect_ClientContext()
        {
            //Act
            ClientContext context = new SharePointServices().GetClientContext();

            Assert.NotNull(context);
            Assert.IsType<ClientContext>(context);
        }
        [Fact]
        public void TryGetSiteSettings_Expect_DTO()
        {
            //Act
            List<SiteSettingsDto> siteSettingsDtos = SharePointFunctions.LoadSiteSettings();
            //Assert
            Assert.NotEmpty(siteSettingsDtos);
            Assert.IsType<List<SiteSettingsDto>>(siteSettingsDtos);
        }
    }
}