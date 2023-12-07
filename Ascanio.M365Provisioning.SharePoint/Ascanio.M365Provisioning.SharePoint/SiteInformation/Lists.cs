using Microsoft.SharePoint.Client;
using ContentType = Microsoft.SharePoint.Client.ContentType;
using List = Microsoft.SharePoint.Client.List;
using Auth0.ManagementApi;
using System.Drawing.Text;
using System.Collections;
using Microsoft.SharePoint.News.DataModel;
using Ascanio.M365Provisioning.SharePoint.Services;

namespace Ascanio.M365Provisioning.SharePoint.SiteInformation
{
    public class Lists
    {
        public Lists()
        {
            SharePointService sharePointService = new();
            using (ClientContext context = sharePointService.GetClientContext())
            {
                ListCollection listCollection = context.Web.Lists;

                context.Load(
                            context.Web.Lists,
                            l => l.Where(l => l.Hidden == false),
                            l => l.Include(
                                                l => l.Title,
                                                l => l.Hidden,
                                                l => l.DefaultDisplayFormUrl,
                                                l => l.BaseType,
                                                l => l.ContentTypes,
                                                l => l.EnableFolderCreation,
                                                l => l.Fields.Include(
                                                                      f => f.InternalName,
                                                                      f => f.Title
                                                                      )
                                               )
                            );
                context.ExecuteQuery();
                foreach (List list in context.Web.Lists)
                {
                    List<string> contentTypes = new();

                    foreach (ContentType contentType in list.ContentTypes)
                    {
                        contentTypes.Add(contentType.Name);
                    }
                    IQueryable<RoleAssignment> queryForList = list.RoleAssignments.Include(roleAsg => roleAsg.Member,
                                                                                           roleAsg => roleAsg.RoleDefinitionBindings.Include(roleDef => roleDef.Name));
                    Dictionary<string, string> listPermissions = GetPermissionDetails(context, queryForList);

                    Guid enterpriseKeywordsValue = Guid.Empty;
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

                    List<ListsDTO> listsDTO = new();
                    listsDTO.Add(new ListsDTO
                                (
                                    list.Title,
                                    list.DefaultViewUrl,
                                    list.BaseType.ToString(),
                                    contentTypes,
                                    list.OnQuickLaunch,
                                    list.EnableFolderCreation,
                                    enterpriseKeywordsValue.ToString(),
                                    true,
                                    listPermissions
                                 ));

                }


            }
            //foreach (List list in web.Lists)
            //{
            //    context.Load(list,
            //        l => l.Title,
            //        l => l.DefaultViewUrl,
            //        l => l.BaseType,
            //        l => l.ContentTypes,
            //        l => l.OnQuickLaunch,
            //        l => l.HasUniqueRoleAssignments,
            //        l => l.Hidden
            //    );
            //    context.Load(list.Fields);
            //    context.ExecuteQuery() ;
            //    // TODO: Hidden test uitvoeren
            //    bool hidden = list.Hidden;
            //    if (!hidden)
            //    {
            //        Guid enterpriseKeywordsValue = Guid.Empty;
            //        try
            //        {
            //            Field enterpriseKeywords = list.Fields.GetByInternalNameOrTitle("TaxKeyword");
            //            context.Load(enterpriseKeywords);
            //            context.ExecuteQuery();
            //            enterpriseKeywordsValue = enterpriseKeywords.Id;
            //        }
            //        catch
            //        {
            //            enterpriseKeywordsValue = Guid.Empty;
            //        }

            //        IQueryable<RoleAssignment> queryForList = list.RoleAssignments.Include(roleAsg => roleAsg.Member,
            //                                                                               roleAsg => roleAsg.RoleDefinitionBindings.Include(roleDef => roleDef.Name));
            //        Dictionary<string, string> listPermissions = GetPermissionDetails(context, queryForList);

                    


            //        lead_ListsDTO.Add(new
            //            (
            //                list.Title,
            //                list.DefaultViewUrl,
            //                list.BaseType.ToString(),
            //                GetContentTypes(context, list),
            //                list.OnQuickLaunch,
            //                GetEnableFolderCreation(context, list),
            //                enterpriseKeywordsValue.ToString(),
            //                list.HasUniqueRoleAssignments,
            //                listPermissions,
            //                hidden
            //            ));
            //    }
            //    }
            //}
            //string jsonFilePath = sharePointService.ListsFilePath ;
            //WriteData2Json writeData2Json = new();
            //writeData2Json.Write2JsonFile(lead_ListsDTO, jsonFilePath);
            //context.Dispose();
        }

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
                permisionDetails.Add(ra.Member.Title, permission);
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
