using Ascanio.M365Provisioning.SharePoint.Services;
using Microsoft.SharePoint.Client;

namespace Ascanio.M365Provisioning.SharePoint
{
    public class Program
    {
        static void Main()
        {
            Run();
        }

        static void Run()
        {
            SharePointService sharePointService = new();
            SharePointFunction sharePointFunction = new();
        }
    }
}