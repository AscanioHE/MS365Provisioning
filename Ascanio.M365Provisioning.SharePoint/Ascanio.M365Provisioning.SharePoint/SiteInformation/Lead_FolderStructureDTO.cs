using Microsoft.SharePoint.Client;

namespace Ascanio.M365Provisioning.SharePoint.SiteInformation
{
    public class Lead_FolderStructureDTO
    {
        public string ListName { get; set; } = string.Empty;
        public string FolderName { get; set; } = string.Empty;
        public List<Lead_FolderStructureDTO> Subfolders { get; set; } = new List<Lead_FolderStructureDTO>();
        public Lead_FolderStructureDTO(string listName, string folderName, List<Lead_FolderStructureDTO> subfolders)
        {
            ListName = listName;
            FolderName = folderName;
            Subfolders = subfolders;
        }

        public Lead_FolderStructureDTO(List<Lead_FolderStructureDTO> subFolders, string folderName, List<Lead_FolderStructureDTO> subfolders)
        {
            FolderName =folderName;
            Subfolders=subfolders;
        }
    }
}
