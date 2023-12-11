using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using M365Provisioning.SharePoint.DTO;


namespace M365Provisioning.SharePoint.Services
{
    public interface ISharePointServices
    {
        (string, string, string, string) GetClientConfiguration();
        ClientContext GetClientContext();
        X509Certificate2 GetCertificateByThumbprint();
        List<SiteSettingsDto> GetSiteSettings();
    }
}
