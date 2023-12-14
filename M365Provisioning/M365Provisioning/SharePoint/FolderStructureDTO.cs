using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;

namespace M365Provisioning.SharePoint
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
