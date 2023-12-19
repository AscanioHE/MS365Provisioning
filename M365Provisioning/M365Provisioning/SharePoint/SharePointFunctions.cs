﻿using System.Collections;
using System.Diagnostics;
using WriteDataToJsonFiles;
using Microsoft.SharePoint.Client;
using System.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Extensions.Logging;
using MS365Provisioning.SharePoint.Model;
using ContentType = Microsoft.SharePoint.Client.ContentType;
using Field = Microsoft.SharePoint.Client.Field;
using NavigationNode = Microsoft.SharePoint.Client.NavigationNode;
using RoleAssignment = Microsoft.SharePoint.Client.RoleAssignment;
using RoleDefinition = Microsoft.SharePoint.Client.RoleDefinition;
using View = Microsoft.SharePoint.Client.View;
using ViewCollection = Microsoft.SharePoint.Client.ViewCollection;
using FieldCollection = Microsoft.SharePoint.Client.FieldCollection;
using List = Microsoft.SharePoint.Client.List;

namespace M365Provisioning.SharePoint
{
    public class SharePointFunctions : ISharePointFunctions
    {
        private ISharePointServices SharePointServices { get; } = new SharePointServices();
        private ClientContext Context { get; set; } = new SharePointServices().Context;

        public ILogger _logger;

        public SharePointFunctions(ILogger logger)
        {
            _logger = logger;
        }

        /*______________________________________________________________________________________________
         Collect Site Settings information
         _______________________________________________________________________________________________*/
        public List<SiteSettingsDto> LoadSiteSettings()
        {
            List<SiteSettingsDto> webTemplatesDto = new();
            Web web = Context.Web;
            Context.Load(web);
            try
            {
                Context.ExecuteQuery();

                WebTemplateCollection webTemplateCollection = web.GetAvailableWebTemplates(1033, true);
                Context.Load(webTemplateCollection);
                Context.ExecuteQuery();


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
                _logger?.LogInformation($"Error executing query: {ex.Message}");
                return new List<SiteSettingsDto>();
            }
            finally
            {
                Context.Dispose();
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
                _logger?.LogInformation($"Error writing data to Json file : {ex.Message}");
            }
            return webTemplatesDto;
        }

        /*______________________________________________________________________________________________
         Collect List Settings information
         _______________________________________________________________________________________________*/
        public List<ListsSettingsDto> LoadListsSettings()
        {
            List<ListsSettingsDto> listDtos = new();
            ListCollection listCollection = Context.Web.Lists;
            Context.Load(Context.Web.Navigation,
                        n => n.QuickLaunch);
            Context.Load(listCollection,
                         lc => lc.Where(l => l.Hidden == false));
            try
            {
                Context.ExecuteQuery();
                foreach (List list in listCollection)
                {
                    Context.Load
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
                    Context.Load(list.Fields);
                    try
                    {
                        Context.ExecuteQuery();
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogInformation($"Error fetching ListSettings : {ex.Message}");
                        
                    }

                    //List<string> contentTypes;
                    //try
                    //{
                    //    contentTypes = GetListContentTypes(list);
                    //}
                    //catch (Exception ex)
                    //{
                    //    _logger?.LogInformation($"Error fetching ContentTypes : {ex.Message}");
                        
                    //}


                    //Dictionary<string, string> listPermissions;
                    try
                    {
                        IQueryable<RoleAssignment> queryForList = list.RoleAssignments.Include(
                            roleAsg => roleAsg.Member,
                            roleAsg => roleAsg.RoleDefinitionBindings.Include(roleDef => roleDef.Name));
                        //listPermissions = GetPermissionDetails(Context, queryForList);
                    }
                    catch (Exception ex)
                    {
                       _logger?.LogInformation($"Error fetching Permissions : {ex.Message}");
                        
                    }

                    //Guid enterpriseKeywordsValue;
                    //try
                    //{
                    //    //enterpriseKeywordsValue = GetEnterpriseKeywordsValue();
                    //}
                    //catch (Exception ex)
                    //{
                    //    _logger?.LogInformation($"Error fetching EnterpriseKeywordsValue : {ex.Message}");
                        
                    //}

                    List<string> quickLaunchHeaders = GetQuickLaunchHeaders();
                    try
                    {
                        //listDtos.Add(new ListsSettingsDto
                        //                (
                        //                    list.Title,
                        //                    list.DefaultViewUrl,
                        //                    list.BaseType.ToString(),
                        //                    contentTypes,
                        //                    list.OnQuickLaunch,
                        //                    quickLaunchHeaders,
                        //                    list.EnableFolderCreation,
                        //                    enterpriseKeywordsValue,
                        //                    // TODO: Unique Role Assignments
                        //                    breakRoleInheritance: true,
                        //                    null
                        //                )
                        //);
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogInformation($"Error writing to DTO File :{ex.Message}");
                        
                    }
                }

                WriteDataToJsonFile(SharePointServices.ListSettingsFilePath, listDtos);
            }
            catch (Exception ex)
            {
                _logger?.LogInformation($"Error fetching ListCollection : {ex.Message}");
                
            }
            finally
            {
                Context.Dispose();
            }

            return listDtos;
        }

        private List<string> GetQuickLaunchHeaders()
        {
            List<string> quickLaunchHeaders = new();
            foreach (NavigationNode navigationNode in Context.Web.Navigation.QuickLaunch)
            {
                Context.Load
                (
                    navigationNode,
                    n => n.Children
                );
                try
                {
                    Context.ExecuteQuery();
                    foreach (NavigationNode childNode in navigationNode.Children)
                    {
                        quickLaunchHeaders.Add(childNode.Title);
                    }
                }
                catch (Exception ex)
                {
                    _logger?.LogInformation($"Error fetching ClientContext: {ex}");
                    
                }
            }

            return quickLaunchHeaders;
        }

        private Guid GetEnterpriseKeywordsValue()
        {
            Guid enterpriseKeywordsValue = Guid.Empty;

            try
            {
                Field enterpriseKeywords = Context.Web.Fields.GetByInternalNameOrTitle("EnterpriseKeywords");

                if (enterpriseKeywords != null)
                {
                    Context.Load(enterpriseKeywords);
                    Context.ExecuteQuery();
                    enterpriseKeywordsValue = enterpriseKeywords.Id;
                }
            }
            catch (Exception ex)
            {
                // Log the exception
                _logger?.LogInformation($"Error fetching Enterprise Keywords value: {ex.Message}");
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
                _logger?.LogInformation($"Error fetching ContentTypes: {ex.Message}");

                // Return an empty list
                contentTypes.Clear();
            }

            return contentTypes;
        }


        Dictionary<string, string> GetPermissionDetails(IQueryable<RoleAssignment> queryString)
        {
            IEnumerable roles = Context.LoadQuery(queryString);
            try
            {
                Context.ExecuteQuery();

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
                _logger?.LogInformation($"Error fetching permissions : {ex}");
                return new Dictionary<string, string>();
            }
        }

        /*______________________________________________________________________________________________
         Collect Listview information
         _______________________________________________________________________________________________*/
        public List<ListViewDto> LoadListViews()
        {
            List<ListViewDto> listViewsDtos = new(); 
            try
            {
                ListCollection listViewslists = Context.Web.Lists;
                Context.Load(listViewslists,
                    lc => lc.Where(
                        l => l.Hidden == false));
                Context.ExecuteQuery();
                WriteDataToJsonFile(SharePointServices.ListViewsFilePath, listViewsDtos);
            }
            catch (Exception ex)
            {
                _logger?.LogInformation($"Error fetching ClientContext : {ex.Message}");
                
            }
            finally
            {
                Context.Dispose();
            }
            return listViewsDtos;
        }

        private List<ListViewDto> GetListViews(List list)
        {
            List<ListViewDto> listViewsDtos = new();
            ViewCollection listViews = list.Views;
            Context.Load(listViews);
            try
            {
                Context.ExecuteQuery();
                foreach (View listView in listViews)
                {
                    try
                    {
                        Context.Load(listView,
                            lv => lv.Title,
                            lv => lv.DefaultView,
                            lv => lv.RowLimit,
                            lv => lv.ViewFields,
                            lv => lv.Scope);
                        Context.ExecuteQuery();

                        listViewsDtos.Add(new ListViewDto(
                            list.Title,listView.Title,listView.DefaultView,new List<string>(),listView.RowLimit,
                            listView.Scope.ToString(),$"{list.Title}.json"));//listView.ViewFields
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogInformation($"Error fetching listview properties : {ex.Message}");
                        
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.LogInformation($"Error fetching Listviews : {ex.Message}");
            }
            return listViewsDtos;
        }

        /*______________________________________________________________________________________________
         Collect SiteColumn information
         _______________________________________________________________________________________________*/
        public List<SiteColumnsDto> LoadSiteColumnsDtos()
        {
            List<SiteColumnsDto> siteColumnsDtos = new();
            ClientContext context = SharePointServices.GetClientContext();
            try
            {
                FieldCollection siteColumns = Context.Web.Fields;
                Context.Load(siteColumns,
                             scc => scc.Include(
                                                                            sc=>sc.Hidden,
                                                                            sc=>sc.InternalName,
                                                                            sc=>sc.SchemaXml,
                                                                            sc=>sc.DefaultValue));
                try
                {
                    Context.ExecuteQuery();
                    foreach (Field siteColumn in siteColumns)
                    {
                        siteColumnsDtos.Add(new SiteColumnsDto(
                            siteColumn.InternalName, siteColumn.SchemaXml, siteColumn.DefaultValue));
                    }

                    WriteDataToJsonFile(SharePointServices.SiteColumnsFilePath, siteColumnsDtos);
                    return siteColumnsDtos;
                }
                catch (Exception ex)
                {
                    _logger?.LogInformation($"Error fetching Site Column settings : {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                _logger?.LogInformation($"Error fetching ContextClient :  {ex.Message}");
            }
            return siteColumnsDtos;
        }
        /*______________________________________________________________________________________________
         Flect contenttypes
         _______________________________________________________________________________________________*/
        public List<ContentTypesDto> LoadContentTypes()
        {
            List<ContentTypesDto> contentTypesDtos = new();
            try
            {
                ListCollection listCollection = Context.Web.Lists;
                Context.Load(listCollection,
                             lc =>lc.Include(
                                                                        l=> l.Hidden == false,
                                                                        l=> l.ContentTypes,
                                                                        l=> l.ContentTypes.Include(
                                                                                                        ct => ct.Name,
                                                                                                        ct=> ct.Parent
                                                                                                        )));
                Context.ExecuteQuery();
                foreach (List list in listCollection)
                {
                    foreach (ContentType contentType in list.ContentTypes)
                    {
                        //ToDo: contenttype required?
                        contentTypesDtos.Add(new(contentType.Name, contentType.Parent.ToString(), "test", false));
                    }
                }
                return contentTypesDtos;
            }
            catch (Exception ex)
            {
                _logger?.LogInformation($"Error fetching ContentTypes : {ex.Message}");
            }
            return contentTypesDtos;
        }
        /*______________________________________________________________________________________________
         Write all data to json file
         _______________________________________________________________________________________________*/
        private void WriteDataToJsonFile(string filePath, object jsonFile)
        {

            WriteDataToJsonFile writeDataToJson = new()
            {
                DtoFile = jsonFile,
                JsonFilePath = filePath
            };
            writeDataToJson.Write2JsonFile();
        }
    }
}
