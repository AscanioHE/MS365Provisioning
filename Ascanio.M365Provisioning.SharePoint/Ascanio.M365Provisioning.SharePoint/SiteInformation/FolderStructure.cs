using Ascanio.M365Provisioning.SharePoint.Services;
using Microsoft.SharePoint.Client;
using System.Collections.Generic;
using System.Linq;

namespace Ascanio.M365Provisioning.SharePoint.SiteInformation
{
    public class FolderStructure
    {
        public List<FolderStructureDto> lead_FolderStructureDTO { get; set; } = new List<FolderStructureDto>();

        public FolderStructure()
        {
            SharePointService sharePointService = new ();
            using (ClientContext context = sharePointService.GetClientContext())
            {
                Web web = context.Web;
                context.Load(
                    web,
                    w => w.Lists
                );
                ListCollection lists = context.Web.Lists;
                context.Load(context.Web.Lists);
                context.ExecuteQuery();
                List<FolderStructureDto> lead_FolderStructureDTOs = new();
                foreach (List list in lists)
                {
                    context.Load
                        (
                        list,
                        l => l.BaseTemplate,
                        l => l.Fields,
                        l => l.Title,
                        l => l.RootFolder.Name,
                        l => l.RootFolder.Folders,
                        l => l.Hidden
                        );
                    context.ExecuteQuery();
                    if(!list.Hidden)
                    {
                        List<Folder> folders = new(list.RootFolder.Folders);
                        foreach (Folder map in folders)
                        {

                            List<FolderStructureDto> subFolders = GetSubFolders(context, map, list);
                            lead_FolderStructureDTO.Add(new FolderStructureDto
                            {
                                ListName = list.Title,
                                FolderName = map.Name,
                                SubFolders = subFolders
                            });
                        }
                    }
                }
                WriteData2Json writeData2Json = new();
                string filePath = sharePointService.FolderStructureFilePath;
                writeData2Json.Write2JsonFile(lead_FolderStructureDTO, filePath);
            }
        }

        private List<FolderStructureDto> GetSubFolders(ClientContext context, Folder folder, List list)
        {
            List<FolderStructureDto> subFolders = new();
            context.Load
                (
                folder, 
                f => f.Folders, 
                f => f.Name
                );
            context.ExecuteQuery();
            if (folder.Folders.Count > 0)
            {
            foreach (Folder subFolder in folder.Folders)
            {
                    subFolders.Add(new FolderStructureDto
                {
                    ListName = list.Title,
                    FolderName = subFolder.Name,
                    SubFolders = GetSubFolders(context, subFolder, list)
                    });
            }
            }
            return subFolders;
        }
    }
}
