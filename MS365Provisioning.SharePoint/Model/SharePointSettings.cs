using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MS365Provisioning.SharePoint.Model
{
    public class SharePointSettings
    {
        public string? ClientId { get; set; }
        public string? TenantId { get; set; }
        public string? ThumbPrint { get; set; }
        public string? SiteUrl { get; set; }
        public string? CientSecret { get; set; }        
    }
    public class FileSettings
    {
        public string? FolderStructureFilePath { get; set; }
        public string? ListsFilePath { get; set; }
        public string? ListViewsFilePath { get; set; }
        public string? SiteColumnsFilePath { get; set; }
        public string? SiteSettingsFilePath { get; set; }
        public string? SitePermissionsFilePath { get; set; }
        public string? WebPartsFilePath { get; set; }
        public string? ContentTypesFilePath { get; set; }

    }
}
