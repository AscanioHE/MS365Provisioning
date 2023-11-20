namespace Ascanio.M365Provisioning.SharePoint
{
    public class ListInfoDTO
    {
        public string Title { get; set; } = string.Empty;
        public string ServerRelativeUrl { get; set; } = string.Empty;
        public int BaseTemplate { get; set; } = 0;
        public List<ContentTypeInfoDTO>? ContentTypes { get; set; }
        public bool OnQuickLaunch { get; set; }
        public bool EnableFolderCreation { get; set; }
        public bool HasUniqueRoleAssignments { get; set; }

        // Add other properties you need
        public ListInfoDTO()
        {
        }

        public ListInfoDTO(string title, string serverRelativeUrl, int baseTemplate, List<ContentTypeInfoDTO>? contentTypes, bool onQuickLaunch, bool enableFolderCreation, bool hasUniqueRoleAssignments)
        {
            Title=title;
            ServerRelativeUrl=serverRelativeUrl;
            BaseTemplate=baseTemplate;
            ContentTypes=contentTypes;
            OnQuickLaunch=onQuickLaunch;
            EnableFolderCreation=enableFolderCreation;
            HasUniqueRoleAssignments=hasUniqueRoleAssignments;
        }



    }
    public class ContentTypeInfoDTO
    {
        public string Name { get; set; } = string.Empty;

        // Add other properties you need for content types
    }
}
