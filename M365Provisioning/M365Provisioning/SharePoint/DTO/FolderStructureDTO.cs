using Microsoft.SharePoint.Client;

namespace Ascanio.M365Provisioning.SharePoint.SiteInformation
{
    public class FolderStructureDTO
    {
        public string ListName { get; set; } = string.Empty;
        public string FolderName { get; set; } = string.Empty;
        public List<FolderStructureDTO> SubFolders { get; set; } = new List<FolderStructureDTO>();
        public FolderStructureDTO(string listName, string folderName, List<FolderStructureDTO> subfolders)
        {
            ListName = listName;
            FolderName = folderName;
            SubFolders = subfolders;
        }
        public FolderStructureDTO() { }
    }
}
