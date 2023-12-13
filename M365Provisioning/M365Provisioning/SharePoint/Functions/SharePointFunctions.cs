using System.Collections;
using System.Diagnostics.Contracts;
using M365Provisioning.SharePoint.Interfaces;
using M365Provisioning.SharePoint.Services;
using M365Provisioning.SharePoint.DTO;

using Microsoft.SharePoint.Client;

namespace M365Provisioning.SharePoint.Functions
{
    public class SharePointFunctions : ISharePointFunctions
    {


        private ISharePointServices SharePointServices { get; set; } = new SharePointServices();

        public List<SiteSettingsDto> LoadSiteSettings()
        {
            string jsonFilePath = SharePointServices.SiteSettingsFilePath;

            List<SiteSettingsDto> webTemplatesDto = new();
            ClientContext context = SharePointServices.GetClientContext();
            Web web = context.Web;
            context.Load(web);
            try
            {
                context.ExecuteQuery();

                WebTemplateCollection webtTemplateCollection = web.GetAvailableWebTemplates(1033, true);
                context.Load(webtTemplateCollection);
                context.ExecuteQuery();


                foreach (WebTemplate template in webtTemplateCollection)
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
                Console.WriteLine($"Error executing query: {ex.Message}");
                return new List<SiteSettingsDto>();
            }
            finally
            {
                context.Dispose();
            }
            WriteDataToJson writeDataToJson = new ();
            writeDataToJson.Write2JsonFile(webTemplatesDto, jsonFilePath);
            return webTemplatesDto;
        }

        public List<ListDto> GetLists()
        {
            List<ListDto> listDtos = new ();
            string jsonFilePath = SharePointServices.ListsFilePath;
            ClientContext context = SharePointServices.GetClientContext();
            ListCollection listCollection = context.Web.Lists;
            context.Load(context.Web.Navigation,
                        n => n.QuickLaunch);
            context.Load(listCollection,
                         lc => lc.Where(
                                                                    l =>l.Hidden == false));
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

                    listDtos.Add(new ListDto
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
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }
            finally
            {
                context.Dispose();
            }

            WriteDataToJson writeDataToJson = new WriteDataToJson();
            writeDataToJson.Write2JsonFile(listDtos, jsonFilePath);
            return listDtos;
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
