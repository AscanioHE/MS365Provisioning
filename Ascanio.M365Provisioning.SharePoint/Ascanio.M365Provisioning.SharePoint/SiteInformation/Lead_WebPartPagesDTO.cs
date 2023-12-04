namespace Ascanio.M365Provisioning.SharePoint.SiteInformation
{
    public class Lead_WebPartPagesDTO
    {
        public string Title { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string QuickLaunchHeader { get; set; } = string.Empty;
        public bool ShowComments { get; set; }
        public string WebPartType { get; set; } = string.Empty;
        public string List { get; set; } = string.Empty;
        public string View { get; set; } = string.Empty;

        public Lead_WebPartPagesDTO(string title, string name, string quickLaunchHeader,
            bool showComments, string webPartType, string list, string view)
        {
            Title = title;
            Name = name;
            QuickLaunchHeader = quickLaunchHeader;
            ShowComments = showComments;
            WebPartType = webPartType;
            List = list;
            View = view;
        }
    }
}
