
using M365Provisioning.SharePoint.Interfaces;
using M365Provisioning.SharePoint.Services;
using Microsoft.SharePoint.Client;

namespace M365Provisioning.Test
{
    public class SharePointServicesTest
    {
        private ISharePointServices _sharePointServices { get; set; }

        public SharePointServicesTest()
        {
            //Arrange
            _sharePointServices = new SharePointServices();
        }

        [Fact]
        public void Try_GetClientContext_Certificate()
        {
            //Act
            ClientContext context = _sharePointServices.GetClientContext();

            Assert.NotNull(context);
            Assert.IsType<ClientContext>(context);
        }
    }
}