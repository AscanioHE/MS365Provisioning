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
    public class Lead_Lists
    {
        public Lead_Lists(ClientContext context, Web web)
        {
            List<Lead_ListsDTO> lead_ListsDTO = new();
            web = context.Web;
            context.Load(web, w => w.Lists);
            context.ExecuteQuery();
            foreach (Microsoft.SharePoint.Client.List list in web.Lists)
            {
                context.Load(list,
                    l => l.Title,
                    l => l.DefaultViewUrl,
                    l => l.BaseType,
                    l => l.ContentTypes,
                    l => l.OnQuickLaunch,
                    l => l.HasUniqueRoleAssignments
                );
                context.Load(list.Fields);
                context.ExecuteQuery() ;

                Guid enterpriseKeywordsValue = Guid.Empty;
                try { 
                    Field enterpriseKeywords = list.Fields.GetByInternalNameOrTitle("taxonomy");
                    context.Load(enterpriseKeywords);
                    context.ExecuteQuery();
                    enterpriseKeywordsValue = enterpriseKeywords.Id;
                }
                catch 
                {
                     enterpriseKeywordsValue = Guid.Empty;
                }
                
                IQueryable<RoleAssignment> queryForList = list.RoleAssignments.Include(roleAsg => roleAsg.Member,
                                                                                       roleAsg => roleAsg.RoleDefinitionBindings.Include(roleDef => roleDef.Name));
                Dictionary<string, string> listPermissions = GetPermissionDetails(context, queryForList);

                
                lead_ListsDTO.Add(new
                    (
                        list.Title,
                        list.DefaultViewUrl,
                        list.BaseType.ToString(),
                        GetContentTypes(context, list),
                        list.OnQuickLaunch,
                        GetEnableFolderCreation(context, list),
                        enterpriseKeywordsValue.ToString(),
                        list.HasUniqueRoleAssignments,
                        listPermissions
                    ));
            }
            string jsonFilePath = "JsonFiles/lead_List.json";
            WriteData2Json writeData2Json = new();
            writeData2Json.Write2JsonFile(lead_ListsDTO, jsonFilePath);
            context.Dispose();
        }

        private Dictionary<string, string> GetPermissionDetails(ClientContext context, IQueryable<RoleAssignment> queryString)
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
