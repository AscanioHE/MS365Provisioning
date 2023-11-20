namespace Ascanio.M365Provisioning.SharePoint
{
    public class DocumentLibraryInfoDTO
    {
        public string Title { get; set; } = string.Empty;
        public string Description {  get; set; } = string.Empty;
        public string DefaultViewUrl {  get; set; } = string.Empty;
        public int ItemCount { get; set; }
        public bool EnableVersioning { get; set; }
        public bool HasUniqueRoleAssignments { get; set; }
        public bool ContentTypesEnabled { get; set; }
        public string ContentType { get; set; }


        public DocumentLibraryInfoDTO() 
        { 
        
        }

        public DocumentLibraryInfoDTO(string title, string description, string defaultViewUrl, int itemCount, bool enableVersioning, bool hasUniqueRoleAssignments, bool contentTypesEnabled)
        {
            Title=title;
            Description=description;
            DefaultViewUrl=defaultViewUrl;
            ItemCount=itemCount;
            EnableVersioning=enableVersioning;
            HasUniqueRoleAssignments=hasUniqueRoleAssignments;
            ContentTypesEnabled=contentTypesEnabled;
        }
    }

}
