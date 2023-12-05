using Microsoft.SharePoint.Client;

namespace Ascanio.M365Provisioning.SharePoint.SiteInformation
{
    public class Lead_FolderStructureDTO
    {
        public string ListName { get; set; } = string.Empty;
        public string FolderName { get; set; } = string.Empty;
        public List<Lead_FolderStructureDTO> SubFolders { get; set; } = new List<Lead_FolderStructureDTO>();
        public Lead_FolderStructureDTO(string listName, string folderName, List<Lead_FolderStructureDTO> subfolders)
        {
            ListName = listName;
            FolderName = folderName;
            SubFolders = subfolders;
        }
        public Lead_FolderStructureDTO() { }
    }
}
