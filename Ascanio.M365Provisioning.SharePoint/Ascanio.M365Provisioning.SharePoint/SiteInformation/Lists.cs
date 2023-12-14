using Microsoft.SharePoint.Client;
using ContentType = Microsoft.SharePoint.Client.ContentType;
using List = Microsoft.SharePoint.Client.List;
using Auth0.ManagementApi;
using System.Drawing.Text;
using System.Collections;
using Microsoft.SharePoint.News.DataModel;
using Ascanio.M365Provisioning.SharePoint.Services;
using AngleSharp.Dom;
using PnP.Core.Model.SharePoint;

namespace Ascanio.M365Provisioning.SharePoint.SiteInformation
{
    public class Lists
    {
        public Lists()
        {
            SharePointService sharePointService = new();
            using ClientContext context = sharePointService.GetClientContext();
            Web web = context.Web;
            List<ListDTO> listsDTO = new();
            web = context.Web;
            context.Load
                (
                web,
                w => w.Lists,
                w => w.Lists.Where(l => l.Hidden == false),
                w => w.Navigation.QuickLaunch
                );
            context.ExecuteQuery();
            lock (web.Lists)
            {
                foreach (List list in web.Lists)
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
                    context.ExecuteQuery();

                    Guid enterpriseKeywordsValue = Guid.Empty;

                    List<string> contentTypes = new();

                    foreach (ContentType contentType in list.ContentTypes)
                    {
                        contentTypes.Add(contentType.Name);
                    }
                    IQueryable<RoleAssignment> queryForList = list.RoleAssignments.Include(roleAsg => roleAsg.Member,
                                                                                                   roleAsg => roleAsg.RoleDefinitionBindings.Include(roleDef => roleDef.Name));
                    Dictionary<string, string> listPermissions = GetPermissionDetails(context, queryForList);

                    try
                    {
                        Field enterpriseKeywords = list.Fields.GetByInternalNameOrTitle("TaxKeyword");
                        context.Load(enterpriseKeywords);
                        context.ExecuteQuery();
                        enterpriseKeywordsValue = enterpriseKeywords.Id;
                    }
                    catch
                    {
                        enterpriseKeywordsValue = Guid.Empty;
                    }
                    List<string> quickLaunchHeaders = new();
                    foreach (NavigationNode navigationNode in context.Web.Navigation.QuickLaunch)
                    {
                        context.Load
                            (
                            navigationNode,
                            n => n.Children
                            );
                        context.ExecuteQuery();
                        foreach (NavigationNode childNode in navigationNode.Children)
                        {
                            quickLaunchHeaders.Add(childNode.Title.ToString());
                        }
                    }

                    listsDTO.Add(new ListDTO
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
                                    true,
                                    listPermissions
                                    )
                    );
                }
            }
            WriteData2Json writeData2Json = new();
            writeData2Json.Write2JsonFile(listsDTO, sharePointService.ListsFilePath);
        }

        //private List<ListsDTO> GetListProperties(ClientContext context, List<ListsDTO> lead_ListsDTO, List propertylist)
        //{
        //    Guid enterpriseKeywordsValue = Guid.Empty;
        //    context.Load(
        //                 context.Web.Lists,
        //                 l => l.Where(l => l.Hidden == false),
        //                 l => l.Include(
        //                                l => l.Title,
        //                                l => l.Hidden,
        //                                l => l.DefaultDisplayFormUrl,
        //                                l => l.BaseType,
        //                                l => l.ContentTypes,
        //                                l => l.EnableFolderCreation,
        //                                l => l.Fields.Include(
        //                                                      f => f.InternalName,
        //                                                      f => f.Title
        //                                                      )
        //                                )
        //                    );

        //    context.Load(context.Web.Navigation.QuickLaunch);
        //    context.ExecuteQuery();
        //    List<string> contentTypes = new();

        //    foreach (ContentType contentType in propertylist.ContentTypes)
        //    {
        //        contentTypes.Add(contentType.Name);
        //    }
        //    IQueryable<RoleAssignment> queryForList = propertylist.RoleAssignments.Include(roleAsg => roleAsg.Member,
        //                                                                                   roleAsg => roleAsg.RoleDefinitionBindings.Include(roleDef => roleDef.Name));
        //    Dictionary<string, string> listPermissions = GetPermissionDetails(context, queryForList);
                
        //    try
        //    {
        //        Field enterpriseKeywords = propertylist.Fields.GetByInternalNameOrTitle("TaxKeyword");
        //        context.Load(enterpriseKeywords);
        //        context.ExecuteQuery();
        //        enterpriseKeywordsValue = enterpriseKeywords.Id;
        //    }
        //    catch
        //    {
        //        enterpriseKeywordsValue = Guid.Empty;
        //    }
        //    List<string> quickLaunchHeaders = new();
        //    foreach (NavigationNode node in context.Web.Navigation.QuickLaunch)
        //    {
        //        context.Load
        //            (
        //            node,
        //            n => n.Children
        //            );
        //        context.ExecuteQuery();
        //        Console.WriteLine(node.Title.ToString());
        //        foreach (NavigationNode childNode in node.Children)
        //        {
        //            Console.WriteLine(childNode.Title.ToString());
        //            quickLaunchHeaders.Add(childNode.Title.ToString());
        //        }
        //    }

        //    List<ListsDTO> listsDTO = new()
        //    {
        //        new ListsDTO
        //                    (
        //                    propertylist.Title,
        //                    propertylist.DefaultViewUrl,
        //                    propertylist.BaseType.ToString(),
        //                    contentTypes,
        //                    propertylist.OnQuickLaunch,
        //                    quickLaunchHeaders,
        //                    propertylist.EnableFolderCreation,
        //                    enterpriseKeywordsValue,
        //                    // TODO: Unique Role Assignments
        //                    true,
        //                    listPermissions
        //                    )
        //    };
        //    return listsDTO;
        //}

        Dictionary<string, string> GetPermissionDetails(ClientContext context, IQueryable<RoleAssignment> queryString)
        {
            IEnumerable roles = context.LoadQuery(queryString);
            context.ExecuteQuery();

            Dictionary<string, string> permisionDetails = new();
            foreach (RoleAssignment ra in roles)
            {
                var rdc = ra.RoleDefinitionBindings;
                string permission = string.Empty;
                foreach (var rdbc in rdc)
                {
                    permission += rdbc.Name.ToString() + ", ";
                }
                permisionDetails.Add(permission, ra.Member.Title);
            }
            return permisionDetails;
        }

        private List<string> GetContentTypes(ClientContext context, List list)
        {
            context.Load(list, l => l.ContentTypes);
            context.ExecuteQuery();

            List<string> contentTypes = new();

            foreach (ContentType contentType in list.ContentTypes)
            {
                contentTypes.Add(contentType.Name);
            }

            return contentTypes;
        }
        private bool GetEnableFolderCreation(ClientContext context, List list)
        {
            context.Load(list, l => l.EnableFolderCreation);
            context.ExecuteQuery();
            return list.EnableFolderCreation;
        }       
    }
}
