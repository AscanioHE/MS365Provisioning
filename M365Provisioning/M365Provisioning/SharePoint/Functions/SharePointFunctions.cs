using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using M365Provisioning.SharePoint.Interfaces;
using PnP.Framework.Provisioning.Model;
using Microsoft.SharePoint.Client;
using M365Provisioning.SharePoint;
using M365Provisioning.SharePoint.Services;
using M365Provisioning.SharePoint.Interfaces;
using M365Provisioning.SharePoint.DTO;

namespace M365Provisioning.SharePoint.Functions
{
    public class SharePointFunctions : ISharePointFunctions
    {
        

        public void SiteSettings()
        {
            ClientContext context = new SharePointServices().Context;
            Web web = context.Web;
            context.Load(web);
            try
            {

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                throw;
            }
            finally
            {
                context.Dispose();
            }
        }

    }
}
