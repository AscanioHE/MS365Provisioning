using System;
using System.Configuration;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Microsoft.SharePoint.Client;

namespace Ascanio.M365Provisioning.SharePoint.Services
{
    public class SharePointService
    {
        public string SiteSettingsFilePath { get; private set; } = string.Empty;
        public string ListsFilePath { get; private set; } = string.Empty;
        public string FolderStructureFilePath { get; private set; } = string.Empty;
        public string ListViewsFilePath { get; private set; } = string.Empty;
        public string SiteColumnsFilePath { get; private set; } = string.Empty;

        public SharePointService()
        {
            LoadFilePathsFromConfiguration();
        }

        private void LoadFilePathsFromConfiguration()
        {
            // Load the configuration file
            var configuration = new ConfigurationBuilder()
                .AddJsonFile("scripts/23M2YZ.Ascanio.AzureFucntions-ApplicationSettings.json", optional: false, reloadOnChange: true)
                .Build();

            SiteSettingsFilePath = configuration["SharePointAscanio:SiteSettingsFilePath"];
            ListsFilePath = configuration["SharePointAscanio:ListsFilePath"];
            FolderStructureFilePath = configuration["SharePointAscanio:FolderStructureFilePath"];
            ListViewsFilePath = configuration["SharePointAscanio:ListViewsFilePath"];
            SiteColumnsFilePath = configuration["SharePointAscanio:SiteColumnsFilePath"];
        }

        public ClientContext GetClientContext()
        {
            (string clientId, string siteUrl, string directoryId, string thumbPrint, string filePath) = GetClientConfiguration();

            // Get the certificate from the local computer with the corresponding thumbprint.
            var certificate = GetCertificateByThumbprint(thumbPrint);

            var authManager = new PnP.Framework.AuthenticationManager(clientId, certificate, directoryId);
            //string fullFilePath = Path.Combine(siteUrl, filePath);

            // Use the PnP Framework to get the SharePoint context.
            ClientContext context = authManager.GetContext(siteUrl);
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

        private static (string clientId, string siteUrl, string directoryId, string thumbPrint, string filePath) GetClientConfiguration()
        {
            // Load the configuration file
            var configuration = new ConfigurationBuilder()
                .AddJsonFile("scripts/23M2YZ.Ascanio.AzureFucntions-ApplicationSettings.json", optional: false, reloadOnChange: true)
                .Build();

            string clientId = configuration["SharePointAscanio:ClientID"];
            string siteUrl = configuration["SharePointAscanio:SiteUrl"];
            string directoryId = configuration["SharePointAscanio:DirectoryId"];
            string thumbPrint = configuration["SharePointAscanio:ThumbPrint"];
            string filePath = configuration["SharePointAscanio:FilePath"];
            return (clientId, siteUrl, directoryId, thumbPrint, filePath);
        }
    }
}
