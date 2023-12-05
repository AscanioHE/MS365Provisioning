using Microsoft.SharePoint.Client;

namespace Ascanio.M365Provisioning.SharePoint.SiteInformation
{
    public class Lead_FolderStructure
    {
        public Lead_FolderStructure(ClientContext context, Web web) 
        {
            context.Load
                (
                web,
                w => w.Lists
                );
            context.ExecuteQuery();
            List<Lead_FolderStructureDTO> lead_FolderStructureDTOs = new ();
            ListCollection lists = context.Web.Lists;
            context.Load(context.Web.Lists);
            context.ExecuteQuery ();
            foreach (List list in lists)
            {
                context.Load (list, l => l.BaseTemplate, l => l.Fields);
                context.ExecuteQuery();
                bool isFolderList = list.BaseTemplate == (int)ListTemplateType.DocumentLibrary && list.Fields.Any(field => field.TypeAsString == "Folder");
                List<Lead_FolderStructureDTO> lead_FolderStructureDTO = new();
                if (isFolderList)
                {
                    if (isFolderList)
                    {
                        var folderStructureDTO = new Lead_FolderStructureDTO
                        {
                            Title = list.Title,
                            Subfolders = GetSubfolders(context, list.RootFolder)
                        };

                        lead_FolderStructureDTOs.Add(folderStructureDTO);

                    }
            }
            
        }

    }
}
