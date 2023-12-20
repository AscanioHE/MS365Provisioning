namespace MS365Provisioning.SharePoint.Model
{
    public class FolderStructureDto
    {
        public string ListName { get; set; } = string.Empty;
        public string FolderName { get; set; } = string.Empty;
        public List<string> SubFolders { get; set; } = new List<string>();
        public FolderStructureDto(string listName, string folderName, List<string> subfolders)
        {
            ListName = listName;
            FolderName = folderName;
            SubFolders = subfolders;
        }
        public FolderStructureDto() { }
    }
}
