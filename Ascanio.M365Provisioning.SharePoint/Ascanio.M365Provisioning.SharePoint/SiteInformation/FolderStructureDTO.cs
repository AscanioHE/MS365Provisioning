using Microsoft.SharePoint.Client;

namespace Ascanio.M365Provisioning.SharePoint.SiteInformation
{
    public class FolderStructureDto
    {
        public string ListName { get; set; } = string.Empty;
        public string FolderName { get; set; } = string.Empty;
        public List<FolderStructureDto> SubFolders { get; set; } = new List<FolderStructureDto>();
        public FolderStructureDto() { }
    }
}
