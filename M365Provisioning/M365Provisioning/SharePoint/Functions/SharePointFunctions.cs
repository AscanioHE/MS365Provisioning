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
            return webTemplatesDto;
        }

        public List<ListDto> GetLists()
        {
            List<ListDto> listDtos = new ();
            List<string> contentTypes = new ();
            string jsonFilePath = SharePointServices.ListsFilePath;
            ClientContext context;
            try
            {
                context = SharePointServices.GetClientContext();
            }
            catch (Exception ex)
            {
                    Console.WriteLine($"Error fetching ClientContext {ex.Message}");
                    throw;
            }
            ListCollection listCollection = context.Web.Lists;
            context.Load(context.Web.Navigation,
                        n => n.QuickLaunch);
            context.Load(listCollection,
                         lc => lc.Where(
                                                                    l =>l.Hidden == false));
            try
            {
                context.ExecuteQuery();
                Dictionary<string, string> listPermissions;
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
                        Console.WriteLine($"Error loading ListSettings : {ex.Message}");
                        throw;
                    }

                    try
                    {
                        contentTypes = GetListContentTypes(list);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error collecting ContentTypes : {ex.Message}");
                        throw;
                    }


                    try
                    {
                        IQueryable<RoleAssignment> queryForList = list.RoleAssignments.Include(
                            roleAsg => roleAsg.Member,
                            roleAsg => roleAsg.RoleDefinitionBindings.Include(roleDef => roleDef.Name));
                        listPermissions = GetPermissionDetails(context, queryForList);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error collecting Permissions : {ex.Message}");
                        throw;
                    }

                    Guid enterpriseKeywordsValue;
                    try
                    {
                        enterpriseKeywordsValue = GetEnterpriseKeywordsValue(list, context);
                    }catch (Exception ex)
                    {
                        Console.WriteLine($"Error collecting EnterpriseKeywordsValue : {ex.Message}");
                    throw;
                    }

                    List<string> quickLaunchHeaders = GetQuickLaunchHeaders(context);

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
                Console.WriteLine($"Error loading ListCollection : {ex.Message}");
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
                        quickLaunchHeaders.Add(childNode.Title.ToString());
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error loading ClientContext: {ex}");
                    throw;
                }
            }

            return quickLaunchHeaders;
        }

        private Guid GetEnterpriseKeywordsValue(List list, ClientContext context)
        {
            Guid enterpriseKeywordsValue;
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

            return enterpriseKeywordsValue;
        }

        private List<string> GetListContentTypes(List list)
        {
            List<string> contentTypes = new();

            foreach (ContentType contentType in list.ContentTypes)
            {
                contentTypes.Add(contentType.Name);
            }

            return contentTypes;
        }

        Dictionary<string, string> GetPermissionDetails(ClientContext context, IQueryable<RoleAssignment> queryString)
        {
            IEnumerable roles = context.LoadQuery(queryString);
            context.ExecuteQuery();

            Dictionary<string, string> permissionDetails = new();
            foreach (RoleAssignment ra in roles)
            {
                RoleDefinitionBindingCollection rdc = ra.RoleDefinitionBindings;
                string permission = string.Empty;
                foreach (RoleDefinition rd in rdc)
                {
                    permission += rd.Name.ToString() + ", ";
                }
                permissionDetails.Add(permission, ra.Member.Title);
            }
            return permissionDetails;
        }
    }
}
