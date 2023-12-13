﻿using Microsoft.SharePoint.Client;
using M365Provisioning.SharePoint.Interfaces;
using System.Security.Cryptography.X509Certificates;
using Microsoft.Extensions.Configuration;

namespace M365Provisioning.SharePoint.Services
{
    public class SharePointServices : ISharePointServices
    {        
        public string SiteSettingsFilePath { get; private set; } 
        public string ListsFilePath { get; private set; }
        public string FolderStructureFilePath { get; private set; } 
        public string ListViewsFilePath { get; private set; }
        public string SiteColumnsFilePath { get; private set; }
        public ClientContext Context { get; private set; }
        public string ClientId { get; private set; }
        public string SiteUrl { get; private set; }
        public string DirectoryId { get; private set; }
        public string ThumbPrint { get; set; }

        public SharePointServices()
        {

            var configuration = new ConfigurationBuilder()
                .AddJsonFile("SharePoint/AppSettings/appsettings.json", optional: false, reloadOnChange: true)
                .Build();

            SiteSettingsFilePath = configuration["SharePoint:SiteSettingsFilePath"]!;
            ListsFilePath = configuration["SharePoint:ListsFilePath"]!;
            FolderStructureFilePath = configuration["SharePoint:FolderStructureFilePath"]!;
            ListViewsFilePath = configuration["SharePoint:ListViewsFilePath"]!;
            SiteColumnsFilePath = configuration["SharePoint:SiteColumnsFilePath"]!;

            ClientId = configuration["SharePointAscanio:ClientID"]!;
            SiteUrl = configuration["SharePointAscanio:SiteUrl"]!;
            DirectoryId = configuration["SharePointAscanio:DirectoryId"]!;
            ThumbPrint = configuration["SharePointAscanio:ThumbPrint"]!;
            X509Certificate2 certificate =  GetCertificateByThumbprint(ThumbPrint);
            var authManager = new PnP.Framework.AuthenticationManager(ClientId, certificate, DirectoryId);
            Context = authManager.GetContext(SiteUrl);
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
    }
}