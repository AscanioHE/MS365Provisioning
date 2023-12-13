using Microsoft.SharePoint.Client;
using M365Provisioning.SharePoint.Services;
using System.Security.Cryptography.X509Certificates;

namespace M365Provisioning.SharePoint.Interfaces;

public interface ISharePointServices
{
    ClientContext GetClientContext();
}

