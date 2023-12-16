using M365Provisioning.SharePoint;
using Microsoft.SharePoint.Client;
using System.Collections.Generic;
using Xunit;

namespace M365Provisioning.Test
{
    public class SharePointFunctionsTest
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
        public void Try_SiteSettings_Expect_DTO()
        {
            ////Act
            //List<SiteSettingsDto> siteSettingsDtos = SharePointFunctions.LoadSiteSettings();
            ////Assert
            //Assert.NotEmpty(siteSettingsDtos);
            //Assert.IsType<List<SiteSettingsDto>>(siteSettingsDtos);
        }

        //[Fact]
        //public void Try_Lists_Expect_DTO()
        //{
        //    //Act
        //    List<ListsSettingsDto> listDtos = SharePointFunctions.LoadListsSettings();
        //    //Assert
        //    Assert.NotEmpty(listDtos);
        //    Assert.IsType<List<ListsSettingsDto>>(listDtos);
        //}

        //[Fact]
        //public void Try_ListViews_Expect_DTO()
        //{
        //    //Act
        //    List<ListViewDto> listViewDtos = SharePointFunctions.LoadListViews();
        //    //Assert
        //    Assert.NotEmpty(listViewDtos);
        //    Assert.IsType<List<ListViewDto>>(listViewDtos);
        //}

        //[Fact]
        //public void Try_SiteColumns_Expect_Dto()
        //{
        //    //Act
        //    List<SiteColumnsDto> siteColumnsDtos = SharePointFunctions.LoadSiteColumnsDtos();
        //    //Assert
        //    Assert.NotEmpty(siteColumnsDtos);
        //    Assert.IsType<List<SiteColumnsDto>>(siteColumnsDtos);

        //}
    }

}