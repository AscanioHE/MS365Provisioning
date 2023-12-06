using Ascanio.M365Provisioning.SharePoint.Services;
using Microsoft.SharePoint.Client;
using System.Collections.Generic;
using System.Linq;

namespace Ascanio.M365Provisioning.SharePoint.SiteInformation
{
    public class Lead_FolderStructure
    {
        public List<Lead_FolderStructureDTO> lead_FolderStructureDTO { get; set; } = new List<Lead_FolderStructureDTO>();

        public Lead_FolderStructure()
        {
            ClientContext context = new SharePointService().GetClientContext();
            Web web = context.Web;
            context.Load(
                web,
                w => w.Lists
            );
            context.ExecuteQuery();

            ListCollection lists = context.Web.Lists;
            context.Load(context.Web.Lists);
            context.ExecuteQuery();
            List<Lead_FolderStructureDTO> lead_FolderStructureDTOs = new();
            foreach (List list in lists)
            {
                context.Load
                    (
                    list, 
                    l => l.BaseTemplate, 
                    l => l.Fields, 
                    l => l.Title, 
                    l => l.RootFolder.Name,
                    l => l.RootFolder.Folders
                    );
                context.ExecuteQuery();
                List<Folder> folders = new(list.RootFolder.Folders);                
                foreach(Folder map in folders)
                {
                    List<Lead_FolderStructureDTO> subFolders = GetSubFolders(context,map,list);
                    lead_FolderStructureDTO.Add(new Lead_FolderStructureDTO
                    {
                        ListName = list.Title,
                        FolderName = map.Name,
                        SubFolders = subFolders
                    });
                }
            }
            WriteData2Json writeData2Json = new();
            string filePath = $"JsonFiles/Lead_FolderStructure.json";
            writeData2Json.Write2JsonFile(lead_FolderStructureDTO, filePath);
            context.Dispose();
        }

        private List<Lead_FolderStructureDTO> GetSubFolders(ClientContext context, Folder folder, List list)
        {
            List<Lead_FolderStructureDTO> subFolders = new();
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
                    subFolders.Add(new Lead_FolderStructureDTO
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
