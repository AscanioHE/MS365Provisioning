using Microsoft.SharePoint.Client;
using M365Provisioning.SharePoint;
using System.Security.Cryptography.X509Certificates;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.Json;
using Microsoft.Graph;
using Newtonsoft.Json;
using System;
using System.Diagnostics;
using Microsoft.Extensions.Logging;

namespace M365Provisioning.SharePoint
{
    public class SharePointServices : ISharePointServices
    {
        public string SiteSettingsFilePath { get; set; }
        public string ListSettingsFilePath { get; set; }
        public string FolderStructureFilePath { get; set; }
        public string ListViewsFilePath { get; set; }
        public string SiteColumnsFilePath { get; set; }
        public ClientContext Context { get; set; }
        public string ClientId { get; set; } = string.Empty;
        public string SiteUrl { get; set; } = string.Empty;
        public string DirectoryId { get; set; } = string.Empty;
        public string ThumbPrint { get; set; } = string.Empty;

        private readonly ILogger _logger = null;

        public SharePointServices()
        {
            try
            {
                ClientContext context = GetClientContext();
                Context = context;
                string appSettingsPath = "appsettings.json";
                IConfigurationRoot configuration = new ConfigurationBuilder()
                    .AddJsonFile(appSettingsPath, optional: false, reloadOnChange: true)
                    .Build();
                SiteSettingsFilePath = configuration["SharePoint:SiteSettingsFilePath"]!;
                ListSettingsFilePath = configuration["SharePoint:ListsFilePath"]!;
                FolderStructureFilePath = configuration["SharePoint:FolderStructureFilePath"]!;
                ListViewsFilePath = configuration["SharePoint:ListViewsFilePath"]!;
                SiteColumnsFilePath = configuration["SharePoint:SiteColumnsFilePath"]!;

            }
            catch (Exception ex)
            {
                _logger?.LogInformation($"Error reading AppSettingsFile : {ex.Message}");
                
            }

        }
        public ClientContext GetClientContext()
        {
            try
            {
                try
                {
                    string appSettingsPath = "appsettings.json";
                    IConfigurationRoot configuration = new ConfigurationBuilder()
                        .AddJsonFile(appSettingsPath, optional: false, reloadOnChange: true)
                        .Build();
                    ClientId = configuration["SharePoint:ClientID"]!;
                    SiteUrl = configuration["SharePoint:SiteUrl"]!;
                    DirectoryId = configuration["SharePoint:DirectoryId"]!;
                    ThumbPrint = configuration["SharePoint:ThumbPrint"]!;
                }
                catch (Exception ex)
                {
                    _logger?.LogInformation($"Error reading AppSetting file : {ex.Message}");
                    return new ClientContext("");
                }


                X509Certificate2 certificate = GetCertificateByThumbprint(ThumbPrint);

                try
                {
                    PnP.Framework.AuthenticationManager authManager = new(ClientId, certificate, DirectoryId);
                    Context = authManager.GetContext(SiteUrl);
                }
                catch (Exception ex)
                {
                    _logger?.LogInformation($"Error creating the ClientContext : {ex.Message}");
                    return new ClientContext("");
                }
                return Context;
            }
            catch (InvalidOperationException ex)
            {
                // Handle the exception here
                _logger?.LogInformation($"Certificate with thumbprint {ThumbPrint} not found!", ex.Message);
                return new ClientContext("");
            }
        }
        private static X509Certificate2 GetCertificateByThumbprint(string thumbprint)
        {
            try
            {
                using X509Store store = new(StoreName.My, StoreLocation.CurrentUser);
                store.Open(OpenFlags.ReadOnly);
                var certificates = store.Certificates.Find(X509FindType.FindByThumbprint, thumbprint, false);
                if (certificates.Count > 0)
                {
                    return certificates[0];
                }
                else
                {
                    throw new InvalidOperationException($"Certificate with thumbprint {thumbprint} not found!");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error creating a Certificate : {ex}");
                return new X509Certificate2();
            }
        }
    }
}