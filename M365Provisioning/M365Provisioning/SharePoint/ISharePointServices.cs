using Microsoft.SharePoint.Client;
using System.Security.Cryptography.X509Certificates;

namespace M365Provisioning.SharePoint;

public interface ISharePointServices
{
    ClientContext GetClientContext();
    string SiteSettingsFilePath { get; set; }
    string ListsFilePath { get; set; }
}

