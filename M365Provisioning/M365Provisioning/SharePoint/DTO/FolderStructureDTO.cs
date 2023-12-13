using Microsoft.SharePoint.Client;

namespace M365Provisioning.SharePoint.DTO
{
    public class FolderStructureDto
    {
        public string ListName { get; set; }
        public string FolderName { get; set; }
        public List<FolderStructureDto> SubFolders { get; set; }
        public FolderStructureDto(string listName, string folderName, List<FolderStructureDto> subfolders)
        {
            ListName = listName;
            FolderName = folderName;
            SubFolders = subfolders;
        }
    }
}
