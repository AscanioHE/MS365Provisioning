using Microsoft.SharePoint.Client;
using M365Provisioning.SharePoint.Interfaces;
using System.Security.Cryptography.X509Certificates;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.Json;
using Microsoft.Graph;
using Newtonsoft.Json;

namespace M365Provisioning.SharePoint.Services
{
    public class SharePointServices : ISharePointServices
    {        
        public string SiteSettingsFilePath { get; private set; } = string.Empty;
        public string ListsFilePath { get; private set; } = string.Empty;
        public string FolderStructureFilePath { get; private set; } = string.Empty;
        public string ListViewsFilePath { get; private set; } = string.Empty;
        public string SiteColumnsFilePath { get; private set; } = string.Empty;
        public ClientContext Context { get; private set; } 
        public string ClientId { get; private set; } = string.Empty;
        public string SiteUrl { get; private set; } = string.Empty;
        public string DirectoryId { get; private set; } = string.Empty;
        public string ThumbPrint { get; set; } = string.Empty;

        public SharePointServices()
        {
            ClientContext context = GetClientContext();
            Context = context;
        }
        public ClientContext GetClientContext()
        {
            try
            {
                string appSettingsPath = "SharePoint/AppSettings/appsettings.json";
                var configuration = new ConfigurationBuilder()
                    .AddJsonFile(appSettingsPath, optional: false, reloadOnChange: true)
                    .Build();
                var configurationSection = configuration.GetSection("SharePoint");
                SiteSettingsFilePath = configuration["SharePoint:SiteSettingsFilePath"];
                ListsFilePath = configuration["SharePoint:ListsFilePath"];
                FolderStructureFilePath = configuration["SharePoint:FolderStructureFilePath"];
                ListViewsFilePath = configuration["SharePoint:ListViewsFilePath"];
                SiteColumnsFilePath = configuration["SharePoint:SiteColumnsFilePath"];


                X509Certificate2 certificate = GetCertificateByThumbprint(ThumbPrint);
                var authManager = new PnP.Framework.AuthenticationManager(ClientId, certificate, DirectoryId);
                Context = authManager.GetContext(SiteUrl);
                return Context;
            }
            catch (InvalidOperationException ex)
            {
                // Handle the exception here
                Console.WriteLine($"Certificate with thumbprint {ThumbPrint} not found!");
                throw;
            }
        }
        public static X509Certificate2 GetCertificateByThumbprint(string thumbprint)
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
    }
}