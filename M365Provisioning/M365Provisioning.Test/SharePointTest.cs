using M365Provisioning.SharePoint.DTO;
using M365Provisioning.SharePoint.Functions;
using M365Provisioning.SharePoint.Interfaces;
using M365Provisioning.SharePoint.Services;
using Microsoft.SharePoint.Client;

namespace M365Provisioning.Test
{
    public class SharePointTest
    {
        private ISharePointFunctions SharePointFunctions { get; set; } = new SharePointFunctions();


        [Fact]
        public void Try_GetClientContext_Expect_ClientContext()
        {
            //Act
            ClientContext context = new SharePointServices().Context;

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

        [Fact]
        public void Try_GetLists_Expect_DTO()
        {
            //Act
            List<ListDto> listDtos = SharePointFunctions.GetLists();
            //Assert
            Assert.NotEmpty(listDtos);
            Assert.IsType<List<ListDto>>(listDtos);
        }
    }
}