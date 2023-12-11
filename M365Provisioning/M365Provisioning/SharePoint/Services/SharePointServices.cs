using System;
using Microsoft.Extensions.Configuration;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using M365Provisioning.SharePoint.DTO;
using M365Provisioning.SharePoint.Services;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using PnP.Framework;

namespace M365Provisioning.SharePoint.Services
{
    public class SharePointServices : ISharePointServices
    {
        public string SiteSettingsFilePath { get; private set; } = string.Empty;
        public string ListsFilePath { get; private set; } = string.Empty;
        public string FolderStructureFilePath { get; private set; } = string.Empty;
        public string ListViewsFilePath { get; private set; } = string.Empty;
        public string SiteColumnsFilePath { get; private set; } = string.Empty;

        public SharePointServices() 
        {
        }

        public void Load()
        {
            _ = GetSiteSettings();
        }

        /*___________________________________________________________________________________________________________________
        
        Get SiteSettings 
        _____________________________________________________________________________________________________________________*/
        public List<SiteSettingsDto>? GetSiteSettings()
        {
            ClientContext context = new SharePointServices().GetClientContext();
            Web web = context.Web;
            context.Load(
            web
            );
            WebTemplateCollection webTemplateCollection = web.GetAvailableWebTemplates(1033, true);
            context.Load(webTemplateCollection);
            try
            {
                // Execute the query to retrieve the data
                context.ExecuteQuery();

                List<SiteSettingsDto> webTemplatesDTO = new();

                foreach (var template in webTemplateCollection)
                {
                    if (!template.IsHidden)
                    {
                        // Create a Lead_SiteSettingsDTO and add it to the list
                        webTemplatesDTO.Add(new SiteSettingsDto
                        {
                            SiteTemplate = template.Name,
                            Value = template.Lcid
                            // Other properties as needed
                        });
                    }
                }
                return webTemplatesDTO;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error executing query: {ex.Message}");
                return null;
            }
            finally
            {
                context.Dispose();
            }

        }

        //___________________________________________________________________________________________________________


        public ClientContext GetClientContext()
        {
            (string clientId, string siteUrl, string directoryId, string thumbPrint) = GetClientConfiguration();
            X509Certificate2 certificate = GetCertificateByThumbprint(thumbPrint);

            var authManager = new AuthenticationManager(clientId, certificate, directoryId);
            ClientContext context = authManager.GetContext(siteUrl);
            return context;
        }
        private X509Certificate2 GetCertificateByThumbprint(string thumbprint)
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

        private (string clientId, string siteUrl, string directoryId, string thumbPrint) GetClientConfiguration()
        {
            // Load the configuration file
            string jsonFilePath = new("SharePoint/Scripts/TestApplicationSettings.json");
            var configuration = new ConfigurationBuilder()
                .AddJsonFile(jsonFilePath, optional: false, reloadOnChange: true)
                .Build();

            // Declare the variables as nullable types
            string clientId = configuration["SharePoint:ClientID"] ?? string.Empty;
            string siteUrl = configuration["SharePoint:SiteUrl"] ?? string.Empty;
            string directoryId = configuration["SharePoint:DirectoryId"] ?? string.Empty;
            string thumbPrint = configuration["SharePoint:ThumbPrint"] ?? string.Empty;

            // Check for null values
            if (clientId == null)
            {
                throw new Exception("The 'clientId' property in the configuration file is null.");
            }

            if (siteUrl == null)
            {
                throw new Exception("The 'siteUrl' property in the configuration file is null.");
            }

            if (directoryId == null)
            {
                throw new Exception("The 'directoryId' property in the configuration file is null.");
            }

            if (thumbPrint == null)
            {
                throw new Exception("The 'thumbPrint' property in the configuration file is null.");
            }

            return (clientId, siteUrl, directoryId, thumbPrint);
        }
        (string, string, string, string) ISharePointServices.GetClientConfiguration()
        {
            throw new NotImplementedException();
        }
        X509Certificate2 ISharePointServices.GetCertificateByThumbprint()
        { 
            throw new NotImplementedException(); 
        }
        ClientContext ISharePointServices.GetClientContext()
        { 
            throw new NotImplementedException(); 
        }
        List<SiteSettingsDto> ISharePointServices.GetSiteSettings()
        {  throw new NotImplementedException(); }
    }
}