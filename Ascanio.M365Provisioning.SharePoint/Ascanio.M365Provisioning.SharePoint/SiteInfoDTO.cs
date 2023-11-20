namespace Ascanio.M365Provisioning.SharePoint
{
    public class SiteInfoDTO
    {
        public SiteInfoDTO() 
        { 
        
        }

        public string Title { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string Url { get; set; } = string.Empty;
        public string ServerRelativeUrl {  get; set; } = string.Empty;
        public DateTime Created { get; set; }
        public DateTime LastModified { get; set; }

        public SiteInfoDTO(string title, string description,string url, string serverRelativeUrl,DateTime created,DateTime lastModified)
        {
            Title = title;
            Description = description;
            Url = url;
            ServerRelativeUrl = serverRelativeUrl;
            Created = created;
            LastModified = lastModified;
        }

    }
}
