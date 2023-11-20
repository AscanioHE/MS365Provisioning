using PnP.Framework;
using Microsoft.Identity.Client;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Microsoft.Graph.Models;
using SPList = Microsoft.SharePoint.Client.List;
using PnP.Core.Model.SharePoint;
using System.Runtime;
using ListTemplateType = Microsoft.SharePoint.Client.ListTemplateType;

namespace Ascanio.M365Provisioning.SharePoint.Services
{
    public class CreateListDT
    {
        readonly SharePointService sharePointService = new();

        public CreateListDT()
        {
            ClientContext clientContext = sharePointService.GetClientContext();

            Web site = GetWebItemParameters(clientContext);
            SiteInfoDTO siteInfoDTO = (SiteInfoDTO)CreateSiteInfoObject(site);
            WriteSite2JsonFile(siteInfoDTO);

            List<SPList> lists = GetListItemParameters(site, clientContext);
            List<ListInfoDTO> listInfoDTO = (List<ListInfoDTO>)CreateDTOLists(lists);
            WriteLists2JsonFile(listInfoDTO);

            List<SPList> documentLibraries = GetDocumentLibraryItemParameters(site, clientContext);
            List<DocumentLibraryInfoDTO> documentLibraryInfoDTOs = (List<DocumentLibraryInfoDTO>)CreateDTOLists(documentLibraries);


            Console.WriteLine("Lists and libraries information has been written to ListsAndLibrariesInfo.json");
        }
        //________________________________________________________________________________________________________________
        // Collect all SharePoint Site information and write the info to a Json file
        //
        private static Web GetWebItemParameters(ClientContext clientContext)
        {
            // Load the current web object
            Web web = clientContext.Web;
            clientContext.Load(web,
                w => w.Title,
                w => w.Description,
                w => w.ServerRelativeUrl,
                w => w.Created,
                w => w.LastItemModifiedDate);
            clientContext.ExecuteQuery();

            // Return the loaded web object
            return (web);
        }

        static object CreateSiteInfoObject(Web web)
        {
            SiteInfoDTO site = new SiteInfoDTO
            {
                Title = web.Title,
                Description = web.Description,
                ServerRelativeUrl = web.ServerRelativeUrl,
                Created = web.Created,
                LastModified = web.LastItemModifiedDate,
            };
            return site;
        }

        private void WriteSite2JsonFile(SiteInfoDTO siteInfoDTO)
        {
            // Create a Json object to store the SharePoint site, lists and document library
            JObject jsonObject = new();
            string jsonString = JsonConvert.SerializeObject(siteInfoDTO, Formatting.Indented);

            // Write JSON to a file
            System.IO.File.WriteAllText("Jsonfiles\\SiteInfo.json", jsonString);

        }
        //________________________________________________________________________________________________________________
        // Collect all SharePoint site Lists information and write the info to a Json file
        //
        private static List<SPList> GetListItemParameters(Web web, ClientContext clientContext)
        {
            ListCollection lists = web.Lists;
            clientContext.Load(lists);
            clientContext.ExecuteQuery();

            List<SPList> result = new List<SPList>();

            foreach (SPList list in lists)
            {
                clientContext.Load(list, t => t.Title);
                clientContext.Load(list.RootFolder, rf => rf.ServerRelativeUrl);
                clientContext.Load(list, bt => bt.BaseTemplate);
                clientContext.Load(list, ct => ct.ContentTypes.Include(ct => ct.Name));
                clientContext.Load(list, oq => oq.OnQuickLaunch);
                clientContext.Load(list, fc => fc.EnableFolderCreation);
                clientContext.Load(list, ur => ur.HasUniqueRoleAssignments);
                clientContext.ExecuteQuery();

                result.Add(list);
            }

            return result;
        }
        static object CreateDTOLists(List<SPList> lists)
        {
            List<ListInfoDTO> listsInfo = new();
            foreach (SPList list in lists)
            {
                ListInfoDTO listInfo = new ListInfoDTO
                {
                    Title = list.Title,
                    ServerRelativeUrl = list.RootFolder.ServerRelativeUrl,
                    BaseTemplate = list.BaseTemplate,
                    ContentTypes = list.ContentTypes.Select(ct => new ContentTypeInfoDTO { Name = ct.Name }).ToList(),
                    OnQuickLaunch = list.OnQuickLaunch,
                    EnableFolderCreation = list.EnableFolderCreation,
                    HasUniqueRoleAssignments    = list.HasUniqueRoleAssignments,
                };
                listsInfo.Add(listInfo);
            }
            return listsInfo;
        }
        private void WriteLists2JsonFile(List<ListInfoDTO> listsInfoDTO)
        {
            // Create a Json object to store the SharePoint site, lists and document library
            JObject jsonObject = new();
            string jsonString = JsonConvert.SerializeObject(listsInfoDTO, Formatting.Indented);

            // Write JSON to a file
            System.IO.File.WriteAllText("Jsonfiles\\ListsAndLibrariesInfo.json", jsonString);
        }
        //________________________________________________________________________________________________________________
        // Collect all SharePoint site Document Library information and write the info to a Json file
        //
        private static List<SPList> GetDocumentLibraryItemParameters(Web web, ClientContext clientContext)
        {
            ListCollection lists = clientContext.Web.Lists;
            clientContext.Load(lists,
                collection => collection.Include(
                    list => list.Title,
                    list => list.Description,
                    list => list.DefaultViewUrl,
                    list => list.ItemCount,
                    list => list.EnableVersioning,
                    list => list.HasUniqueRoleAssignments,
                    list => list.ContentTypesEnabled,
                    list => list.ContentTypes,
                    list => list.DefaultView
                    )
                );
            clientContext.ExecuteQuery();

            List<SPList> documentLibraries = (List<SPList>)lists.Where(list => list.BaseTemplate == (int)ListTemplateType.DocumentLibrary);

            return documentLibraries;
        }

        // Methode om het lijsttype op basis van het sjabloonnummer op te halen
        static string GetListType(int templateType)
        {
            switch (templateType)
            {
                case (int)ListTemplateType.DocumentLibrary:
                    return "Document Library";
                case (int)ListTemplateType.GenericList:
                    return "Generic List";
                // Voeg andere lijsttypen toe zoals nodig
                default:
                    return "Unknown Type";
            }
        }
        static object CreateDTOcumentLibraries(List<SPList> documentLibraries)
        {
            List<DocumentLibraryInfoDTO> documentLibrariesInfo = new();
            foreach (SPList documentLibrary in documentLibraries)
            {
                documentLibrariesInfo.Add(new DocumentLibraryInfoDTO
                {
                    Title = documentLibrary.Title,
                    Description = documentLibrary.Description,
                    DefaultViewUrl = documentLibrary.DefaultViewUrl,
                    ItemCount = documentLibrary.ItemCount,
                    EnableVersioning = documentLibrary.EnableVersioning,
                    HasUniqueRoleAssignments = documentLibrary.HasUniqueRoleAssignments,
                    ContentTypesEnabled = documentLibrary.ContentTypesEnabled
                });
                // Voeg andere eigenschappen toe indien nodig
            }
            return documentLibrariesInfo;
        }
    }
}
