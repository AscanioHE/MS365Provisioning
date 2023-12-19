namespace MS365Provisioning.SharePoint.Model
{
    public class LeadWebPartPagesDto
    {
        public string Title { get; set; }
        public string Name { get; set; }
        public string QuickLaunchHeader { get; set; }
        public bool ShowComments { get; set; }
        public string WebPartType { get; set; }
        public string List { get; set; }
        public string View { get; set; }

        public LeadWebPartPagesDto(string title, string name, string quickLaunchHeader,
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
