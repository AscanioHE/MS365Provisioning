using System.Collections;
using System.Diagnostics;
using WriteDataToJsonFiles;
using Microsoft.SharePoint.Client;
using System.Text;
using M365Provisioning.SharePoint;
using M365Provisioning.SharePoint.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using PnP.Framework.Provisioning.Model;
using ContentType = Microsoft.SharePoint.Client.ContentType;
using Field = Microsoft.SharePoint.Client.Field;
using NavigationNode = Microsoft.SharePoint.Client.NavigationNode;
using RoleAssignment = Microsoft.SharePoint.Client.RoleAssignment;
using RoleDefinition = Microsoft.SharePoint.Client.RoleDefinition;
using View = Microsoft.SharePoint.Client.View;
using ViewCollection = Microsoft.SharePoint.Client.ViewCollection;

namespace M365Provisioning.SharePoint
{
    public class SharePointFunctions : ISharePointFunctions
    {
        private ISharePointServices SharePointServices { get; } = new SharePointServices();

        /*______________________________________________________________________________________________
         Collect Site Settings information
         _______________________________________________________________________________________________*/
        public List<SiteSettingsDto> LoadSiteSettings()
        {
            List<SiteSettingsDto> webTemplatesDto = new();
            ClientContext context = SharePointServices.GetClientContext();
            Web web = context.Web;
            context.Load(web);
            try
            {
                context.ExecuteQuery();

                WebTemplateCollection webTemplateCollection = web.GetAvailableWebTemplates(1033, true);
                context.Load(webTemplateCollection);
                context.ExecuteQuery();


                foreach (WebTemplate template in webTemplateCollection)
                {
                    webTemplatesDto.Add(new SiteSettingsDto
                    {
                        SiteTemplate = template.Name,
                        Value = template.Lcid
                    });
                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error executing query: {ex.Message}");
                return new List<SiteSettingsDto>();
            }
            finally
            {
                context.Dispose();
            }

            try
            {

                WriteDataToJsonFile writeDataToJson = new()
                {
                    DtoFile = webTemplatesDto,
                    JsonFilePath = SharePointServices.SiteSettingsFilePath
                };
                writeDataToJson.Write2JsonFile();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error writing data to Json file : {ex.Message}");
            }
            return webTemplatesDto;
        }

        /*______________________________________________________________________________________________
         Collect List Settings information
         _______________________________________________________________________________________________*/
        public List<ListsSettingsDto> LoadListsSettings()
        {
            List<ListsSettingsDto> listDtos = new();
            ClientContext context;
            try
            {
                context = SharePointServices.GetClientContext();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error fetching ClientContext {ex.Message}");
                throw;
            }
            ListCollection listCollection = context.Web.Lists;
            context.Load(context.Web.Navigation,
                        n => n.QuickLaunch);
            context.Load(listCollection,
                         lc => lc.Where(
                                                                    l => l.Hidden == false));
            try
            {
                context.ExecuteQuery();
                foreach (List list in listCollection)
                {
                    context.Load
                    (
                        list,
                        l => l.Title,
                        l => l.DefaultViewUrl,
                        l => l.BaseType,
                        l => l.ContentTypes,
                        l => l.OnQuickLaunch,
                        l => l.HasUniqueRoleAssignments,
                        l => l.Fields.Include(
                            f => f.InternalName,
                            f => f.Title
                        )
                    );
                    context.Load(list.Fields);
                    try
                    {
                        context.ExecuteQuery();
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Error loading ListSettings : {ex.Message}");
                        throw;
                    }

                    List<string> contentTypes;
                    try
                    {
                        contentTypes = GetListContentTypes(list);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Error collecting ContentTypes : {ex.Message}");
                        throw;
                    }


                    Dictionary<string, string> listPermissions;
                    try
                    {
                        IQueryable<RoleAssignment> queryForList = list.RoleAssignments.Include(
                            roleAsg => roleAsg.Member,
                            roleAsg => roleAsg.RoleDefinitionBindings.Include(roleDef => roleDef.Name));
                        listPermissions = GetPermissionDetails(context, queryForList);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Error collecting Permissions : {ex.Message}");
                        throw;
                    }

                    Guid enterpriseKeywordsValue;
                    try
                    {
                        enterpriseKeywordsValue = GetEnterpriseKeywordsValue(context);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Error collecting EnterpriseKeywordsValue : {ex.Message}");
                        throw;
                    }

                    List<string> quickLaunchHeaders = GetQuickLaunchHeaders(context);
                    try
                    {
                        listDtos.Add(new ListsSettingsDto
                                        (
                                            list.Title,
                                            list.DefaultViewUrl,
                                            list.BaseType.ToString(),
                                            contentTypes,
                                            list.OnQuickLaunch,
                                            quickLaunchHeaders,
                                            list.EnableFolderCreation,
                                            enterpriseKeywordsValue,
                                            // TODO: Unique Role Assignments
                                            breakRoleInheritance: true,
                                            listPermissions
                                        )
                        );
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Error writing to DTO File :{ex.Message}");
                        throw;
                    }
                }

                WriteDataToJsonFile writeDataToJson = new()
                {
                    DtoFile = listDtos,
                    JsonFilePath = SharePointServices.ListsFilePath
                };
                writeDataToJson.Write2JsonFile();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading ListCollection : {ex.Message}");
                throw;
            }
            finally
            {
                context.Dispose();
            }

            return listDtos;
        }

        private List<string> GetQuickLaunchHeaders(ClientContext context)
        {
            List<string> quickLaunchHeaders = new();
            foreach (NavigationNode navigationNode in context.Web.Navigation.QuickLaunch)
            {
                context.Load
                (
                    navigationNode,
                    n => n.Children
                );
                try
                {
                    context.ExecuteQuery();
                    foreach (NavigationNode childNode in navigationNode.Children)
                    {
                        quickLaunchHeaders.Add(childNode.Title);
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error loading ClientContext: {ex}");
                    throw;
                }
            }

            return quickLaunchHeaders;
        }

        private Guid GetEnterpriseKeywordsValue(ClientContext context)
        {
            Guid enterpriseKeywordsValue = Guid.Empty;

            try
            {
                Field enterpriseKeywords = context.Web.Fields.GetByInternalNameOrTitle("EnterpriseKeywords");

                if (enterpriseKeywords != null)
                {
                    context.Load(enterpriseKeywords);
                    context.ExecuteQuery();
                    enterpriseKeywordsValue = enterpriseKeywords.Id;
                }
            }
            catch (Exception ex)
            {
                // Log the exception
                Debug.WriteLine($"Error retrieving Enterprise Keywords value: {ex.Message}");
            }

            return enterpriseKeywordsValue;
        }

        private List<string> GetListContentTypes(List list)
        {
            List<string> contentTypes = new();

            try
            {
                if (list.ContentTypes.Count == 0)
                {
                    return contentTypes; // No ContentTypes to return
                }

                foreach (ContentType contentType in list.ContentTypes)
                {
                    contentTypes.Add(contentType.Name);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error collecting ContentTypes: {ex.Message}");

                // Return an empty list
                contentTypes.Clear();
            }

            return contentTypes;
        }


        Dictionary<string, string> GetPermissionDetails(ClientContext context, IQueryable<RoleAssignment> queryString)
        {
            IEnumerable roles = context.LoadQuery(queryString);
            try
            {
                context.ExecuteQuery();

                Dictionary<string, string> permissionDetails = new();
                foreach (RoleAssignment ra in roles)
                {
                    RoleDefinitionBindingCollection rdc = ra.RoleDefinitionBindings;
                    StringBuilder permissionBuilder = new();
                    foreach (RoleDefinition rd in rdc)
                    {
                        permissionBuilder.Append(rd.Name + ", ");
                    }
                    string permission = permissionBuilder.ToString();
                    permissionBuilder.Clear();

                    permissionDetails.Add(permission, ra.Member.Title);
                }
                return permissionDetails;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error collecting permissions : {ex}");
                return new Dictionary<string, string>();
            }
        }

        /*______________________________________________________________________________________________
         Collect Listview information
         _______________________________________________________________________________________________*/
        public List<ListViewDto> LoadListViews()
        {
            List<ListViewDto> listViewDtos = new(); 
            ClientContext context = new SharePointServices().GetClientContext();
            try
            {
                ListCollection listViewlists = context.Web.Lists;
                context.Load(listViewlists,
                    lc => lc.Where(
                        l => l.Hidden == false));
                context.ExecuteQuery();
                foreach (List list in listViewlists)
                {
                    
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading ClientContext : {ex.Message}");
                throw;
            }
            finally
            {
                context.Dispose();
            }
            return listViewDtos;
        }

        private List<ListViewDto> GetListViews(ClientContext context, List list)
        {
            List<ListViewDto> listViewDtos = new();
            ViewCollection listViews = list.Views;
            context.Load(listViews);
            try
            {
                foreach (View listView in listViews)
                {
                    context.Load(listView,
                                    lv => lv.Title,
                                    lv =>lv.DefaultView,
                                    lv=>lv.RowLimit,
                                    lv => lv.RowLimit,
                                    lv=> lv.Scope);
                }
                return listViewDtos;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading Listviews : {ex.Message}");
                throw;
            }
        }
    }
}
