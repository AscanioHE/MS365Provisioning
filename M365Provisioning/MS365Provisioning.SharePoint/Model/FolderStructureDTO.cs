namespace MS365Provisioning.SharePoint.Model
{
    public class FolderStructureDto
    {
        public string ListName { get; set; } = string.Empty;
        public string FolderName { get; set; } = string.Empty;
        public List<FolderStructureDto> SubFolders { get; set; } = new List<FolderStructureDto>();
        public FolderStructureDto(string listName, string folderName, List<FolderStructureDto> subfolders)
        {
            ListName = listName;
            FolderName = folderName;
            SubFolders = subfolders;
        }
        public FolderStructureDto() { }
    }
}
