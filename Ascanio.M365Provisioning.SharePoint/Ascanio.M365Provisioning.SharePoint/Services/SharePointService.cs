using System.Net;
using PnP.Framework;
using Microsoft.SharePoint.Client;
using System.IO;
using File = System.IO.File;
using System.Security.Cryptography.X509Certificates;
using PnP.Framework.Modernization.Cache;
using Newtonsoft.Json;
using Microsoft.Identity.Client;
using SPClient = Microsoft.SharePoint.Client;
using Microsoft.Graph.Models;

namespace Ascanio.M365Provisioning.SharePoint.Services
{
    public class SharePointService
    {
        public ClientContext GetClientContext()
        {
            (string clientId, string siteUrl, string directoryId, string thumbPrint) = GetClientConfiguration();

            // Get the certificate from the local computer with the corresponding thumbprint.
            var certificate = GetCertificateByThumbprint(thumbPrint);

            var authManager = new PnP.Framework.AuthenticationManager(clientId, certificate, directoryId);

            // Use the PnP Framework to get the SharePoint context.
            SPClient.ClientContext context = authManager.GetContext(siteUrl);
            return context;
        }    

            private static X509Certificate2 GetCertificateByThumbprint(string thumbprint)
            {
            using X509Store store = new(StoreName.My, StoreLocation.CurrentUser);
            store.Open(OpenFlags.ReadOnly);
            var certificates = store.Certificates.Find(X509FindType.FindByThumbprint, thumbprint, false);
            if (certificates.Count > 0)
            {
                Console.WriteLine("Authenticated and connected to SharePoint!");
                return certificates[0];
            }
            else
            {
                throw new InvalidOperationException($"Certificate with thumbprint {thumbprint} not found!");
            }
        }
        private static (string clientId, string siteUrl, string directoryId, string thumbPrint) GetClientConfiguration()
            {
                // Load the configuration file
                var configuration = new ConfigurationBuilder()
                    .AddJsonFile("Scripts\\23M2YZ.Ascanio.AzureFucntions-ApplicationSettings.json", optional: false, reloadOnChange: true)
                    .Build();

                string clientId = configuration["SharePointAscanio:ClientID"];
                string siteUrl = configuration["SharePointAscanio:SiteUrl"];
                string directoryId = configuration["SharePointAscanio:DirectoryId"];
                string thumbPrint = configuration["SharePointAscanio:ThumbPrint"];

                return (clientId, siteUrl, directoryId, thumbPrint);
            }

    }
}
