using Microsoft.Graph.Models;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using System.Web;
using File = Microsoft.SharePoint.Client.File;
using Folder = Microsoft.SharePoint.Client.Folder;
using List = Microsoft.SharePoint.Client.List;
using ListItem = Microsoft.SharePoint.Client.ListItem;

namespace Ascanio.M365Provisioning.SharePoint.SiteInformation
{
    public class Lead_WebPartPages
    {
        public Lead_WebPartPages(ClientContext context,Web web)
        {
            web = context.Web;
            List<Lead_WebPartPagesDTO> lead_WebPartPagesDTOs = GetWebPartPages(context, web);
            context.Dispose ();
        }
        private List<Lead_WebPartPagesDTO> GetWebPartPages(ClientContext context, Web web)
        {
            return WebPartItems(context, web);
            
            static List<Lead_WebPartPagesDTO> WebPartItems(ClientContext context, Web web)
            {
                context.Load
                (
                    web,
                    w => w.Lists
                );
                context.ExecuteQuery();
                List<Lead_WebPartPagesDTO> lead_WebPartPagesDTO = new();

                foreach (List list in web.Lists)
                {
                    CamlQuery query = CamlQuery.CreateAllItemsQuery();
                    ListItemCollection items = list.GetItems(query);
                    context.Load(items);
                    context.ExecuteQuery();

                    foreach (ListItem item in items)
                    {
                        if (item != null)
                        {
                            context.Load(item,
                                i => i.FileSystemObjectType,
                                i => i.File,
                                i => i.File.ServerRelativeUrl,
                                i => i.DisplayName
                            ) ;
                            context.ExecuteQuery();
                            LimitedWebPartManager? webPartManager = null;
                            switch (item.FileSystemObjectType)
                            {
                                case FileSystemObjectType.File:
                                    {
                                        webPartManager=GetFileInformation(context, item, webPartManager);
                                        if (webPartManager != null)
                                        {
                                            foreach (WebPartDefinition webPartDefinition in webPartManager.WebParts)
                                            {
                                                string? title = webPartDefinition.WebPart.Properties.FieldValues["Title"]?.ToString();
                                                Console.WriteLine(title);
                                            }
                                        }

                                        break;
                                    }

                                case FileSystemObjectType.Folder:
                                    {
                                        GetFolderInformation(context, item, webPartManager);

                                        break;
                                    }
                            }
                        }
                    }
                }
                return lead_WebPartPagesDTO;

                LimitedWebPartManager? GetFileInformation(ClientContext context, ListItem item, LimitedWebPartManager? webPartManager)
                {
                    if ((item["File_x0020_Type"] as string) == "html" || (item["File_x0020_Type"] as string) == "aspx")
                    {
                        try
                        {
                            // Verkrijg de LimitedWebPartManager
                            ListItem fileProperties = item.File.ListItemAllFields;
                            context.Load(fileProperties);
                            context.ExecuteQuery();
                            // Toon de eigenschappen van het bestand
                            foreach (var property in fileProperties.FieldValues)
                            {
                                Console.WriteLine($"{property.Key}: {property.Value}");
                            }
                        }
                        catch (ServerException ex)
                        {
                            // Mogelijk geen webonderdelen op het bestand, verdergaan zonder fouten
                            Console.WriteLine($"Fout bij het ophalen van LimitedWebPartManager. Details: {ex.Message}");
                        }
                    }
                    else
                    {
                        Console.WriteLine("Dit bestand ondersteunt geen webonderdelen.");
                    }

                    return webPartManager;
                }

                void GetFolderInformation(ClientContext context, ListItem item, LimitedWebPartManager? webPartManager)
                {
                    string folderServerRelativeUrl = item["FileDirRef"]?.ToString();

                    
                        Folder folder = context.Web.GetFolderByServerRelativeUrl(folderServerRelativeUrl);
                        context.Load(folder);
                        context.Load(folder.Folders);
                        context.Load(folder.Files);
                        context.ExecuteQuery();

                        // Toon de naam van de map
                        Console.WriteLine($"Mapnaam: {folder.Name}");

                        // Toon de inhoud van de map (bestanden)
                        foreach (File file in folder.Files)
                        {
                            Console.WriteLine($"Bestandsnaam: {file.Name}");
                        }

                        // Toon de inhoud van de map (submappen)
                        foreach (Folder subFolder in folder.Folders)
                        {
                            Console.WriteLine($"Submapnaam: {subFolder.Name}");
                            context.Load(subFolder.Files);
                            context.ExecuteQuery();
                            foreach (File file in subFolder.Files)
                            {
                                Console.WriteLine($"Bestandsnaam: {file.Name}");
                                GetFileInformation(context, item, webPartManager);
                            }
                        }

                }
            }
        }
    }
}
