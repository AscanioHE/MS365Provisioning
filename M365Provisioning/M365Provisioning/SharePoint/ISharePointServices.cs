using Microsoft.SharePoint.Client;
using System.Security.Cryptography.X509Certificates;

namespace M365Provisioning.SharePoint;

public interface ISharePointServices
{
    ClientContext GetClientContext();
    string SiteSettingsFilePath { get; set; }
    string ListSettingsFilePath { get; set; }
    string FolderStructureFilePath { get; set; }
    string ListViewsFilePath { get; set; }
    string SiteColumnsFilePath { get; set; }
    ClientContext Context { get; set; }
    string ClientId { get; set; }
    string SiteUrl { get; set; }
    string DirectoryId { get; set; }
    string ThumbPrint { get; set; }
}

