using AngleSharp.Css.Dom;
using Ascanio.M365Provisioning.SharePoint.Services;
using Microsoft.Graph.Models;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.News.DataModel;
using PnP.Core.Model.SharePoint;
using System.Reflection;
using List = Microsoft.SharePoint.Client.List;
using ListItem = Microsoft.SharePoint.Client.ListItem;
using ListTemplateType = Microsoft.SharePoint.Client.ListTemplateType;

namespace Ascanio.M365Provisioning.SharePoint
{
    public class temp
    {
        public void Test()
        {
            SharePointService sharePointService = new();
            ClientContext context = sharePointService.GetClientContext();
            Web web = context.Web;
            context.ExecuteQuery();
            IEnumerable<List> Libraries = context.LoadQuery(context.Web.Lists.Where(l => l.BaseTemplate == 101));
            context.ExecuteQuery();
            foreach (List lib in Libraries)
            {
                Console.WriteLine(lib.Title);
                GetListProperties(context, lib);


            }


            context.Dispose();
            Console.WriteLine("Einde script");
        }

        private void IdentifyLists(ClientContext context, ListCollection lists)
        {
            foreach (List list in lists)
            {
                switch (list.BaseTemplate)
                {
                    case -1:
                        Console.WriteLine($"{list.Title} | InvalidType: {list.BaseTemplate}");
                        break;
                    case 0:
                        Console.WriteLine($"{list.Title} | NoListTemplate: {list.BaseTemplate}");
                        break;
                    case 100:
                        Console.WriteLine($"{list.Title} | GenericList: {list.BaseTemplate}");
                        break;
                    case 101:
                        Console.WriteLine($"{list.Title} | DocumentLibrary: {list.BaseTemplate}");
                        break;
                    case 102:
                        Console.WriteLine($"{list.Title} | Survey: {list.BaseTemplate}");
                        break;
                    case 103:
                        Console.WriteLine($"{list.Title} | Links: {list.BaseTemplate}");
                        break;
                    case 104:
                        Console.WriteLine($"{list.Title} | Announcements: {list.BaseTemplate}");
                        break;
                    case 105:
                        Console.WriteLine($"{list.Title} | Contacts: {list.BaseTemplate}");
                        break;
                    case 106:
                        Console.WriteLine($"{list.Title} | Events: {list.BaseTemplate}");
                        break;
                    case 107:
                        Console.WriteLine($"{list.Title} | Tasks: {list.BaseTemplate}");
                        break;
                    case 108:
                        Console.WriteLine($"{list.Title} | DiscussionBoard: {list.BaseTemplate}");
                        break;
                    case 109:
                        Console.WriteLine($"{list.Title} | PictureLibrary: {list.BaseTemplate}");
                        break;
                    case 110:
                        Console.WriteLine($"{list.Title} | DataSources: {list.BaseTemplate}");
                        break;
                    case 111:
                        Console.WriteLine($"{list.Title} | WebTemplateCatalog: {list.BaseTemplate}");
                        break;
                    case 112:
                        Console.WriteLine($"{list.Title} | UserInformation: {list.BaseTemplate}");
                        break;
                    case 113:
                        Console.WriteLine($"{list.Title} | WebPartCatalog: {list.BaseTemplate}");
                        break;
                    case 114:
                        Console.WriteLine($"{list.Title} | ListTemplateCatalog: {list.BaseTemplate}");
                        break;
                    case 115:
                        Console.WriteLine($"{list.Title} | XMLForm: {list.BaseTemplate}");
                        break;
                    case 116:
                        Console.WriteLine($"{list.Title} | MasterPageCatalog: {list.BaseTemplate}");
                        break;
                    case 117:
                        Console.WriteLine($"{list.Title} | NoCodeWorkflows: {list.BaseTemplate}");
                        break;
                    case 118:
                        Console.WriteLine($"{list.Title} | WorkflowProcess: {list.BaseTemplate}");
                        break;
                    case 119:
                        Console.WriteLine($"{list.Title} | WebPageLibrary: {list.BaseTemplate}");
                        break;
                    case 120:
                        Console.WriteLine($"{list.Title} | CustomGrid: {list.BaseTemplate}");
                        break;
                    case 121:
                        Console.WriteLine($"{list.Title} | SolutionCatalog: {list.BaseTemplate}");
                        break;
                    case 122:
                        Console.WriteLine($"{list.Title} | NoCodePublic: {list.BaseTemplate}");
                        break;
                    case 123:
                        Console.WriteLine($"{list.Title} | ThemeCatalog: {list.BaseTemplate}");
                        break;
                    case 124:
                        Console.WriteLine($"{list.Title} | DesignCatalog: {list.BaseTemplate}");
                        break;
                    case 125:
                        Console.WriteLine($"{list.Title} | AppDataCatalog: {list.BaseTemplate}");
                        break;
                    case 130:
                        Console.WriteLine($"{list.Title} | DataConnectionLibrary: {list.BaseTemplate}");
                        break;
                    case 140:
                        Console.WriteLine($"{list.Title} | WorkflowHistory: {list.BaseTemplate}");
                        break;
                    case 150:
                        Console.WriteLine($"{list.Title} | GanttTasks: {list.BaseTemplate}");
                        break;
                    case 151:
                        Console.WriteLine($"{list.Title} | HelpLibrary: {list.BaseTemplate}");
                        break;
                    case 160:
                        Console.WriteLine($"{list.Title} | AccessRequest: {list.BaseTemplate}");
                        break;
                    case 171:
                        Console.WriteLine($"{list.Title} | TasksWithTimelineAndHierarchy: {list.BaseTemplate}");
                        break;
                    case 175:
                        Console.WriteLine($"{list.Title} | MaintenanceLogs: {list.BaseTemplate}");
                        break;
                    case 200:
                        Console.WriteLine($"{list.Title} | Meetings: {list.BaseTemplate}");
                        break;
                    case 201:
                        Console.WriteLine($"{list.Title} | Agenda: {list.BaseTemplate}");
                        break;
                    case 202:
                        Console.WriteLine($"{list.Title} | MeetingUser: {list.BaseTemplate}");
                        break;
                    case 204:
                        Console.WriteLine($"{list.Title} | Decision: {list.BaseTemplate}");
                        break;
                    case 207:
                        Console.WriteLine($"{list.Title} | MeetingObjective: {list.BaseTemplate}");
                        break;
                    case 210:
                        Console.WriteLine($"{list.Title} | TextBox: {list.BaseTemplate}");
                        break;
                    case 211:
                        Console.WriteLine($"{list.Title} | ThingsToBring: {list.BaseTemplate}");
                        break;
                    case 212:
                        Console.WriteLine($"{list.Title} | HomePageLibrary: {list.BaseTemplate}");
                        break;
                    case 301:
                        Console.WriteLine($"{list.Title} | Posts: {list.BaseTemplate}");
                        break;
                    case 302:
                        Console.WriteLine($"{list.Title} | Comments: {list.BaseTemplate}");
                        break;
                    case 303:
                        Console.WriteLine($"{list.Title} | Categories: {list.BaseTemplate}");
                        break;
                    case 402:
                        Console.WriteLine($"{list.Title} | Facility: {list.BaseTemplate}");
                        break;
                    case 403:
                        Console.WriteLine($"{list.Title} | Whereabouts: {list.BaseTemplate}");
                        break;
                    case 404:
                        Console.WriteLine($"{list.Title} | CallTrack: {list.BaseTemplate}");
                        break;
                    case 405:
                        Console.WriteLine($"{list.Title} | Circulation: {list.BaseTemplate}");
                        break;
                    case 420:
                        Console.WriteLine($"{list.Title} | Timecard: {list.BaseTemplate}");
                        break;
                    case 421:
                        Console.WriteLine($"{list.Title} | Holidays: {list.BaseTemplate}");
                        break;
                    case 499:
                        Console.WriteLine($"{list.Title} | IMEDic: {list.BaseTemplate}");
                        break;
                    case 600:
                        Console.WriteLine($"{list.Title} | ExternalList: {list.BaseTemplate}");
                        break;
                    case 700:
                        Console.WriteLine($"{list.Title} | MySiteDocumentLibrary: {list.BaseTemplate}");
                        break;
                    case 1100:
                        Console.WriteLine($"{list.Title} | IssueTracking: {list.BaseTemplate}");
                        break;
                    case 1200:
                        Console.WriteLine($"{list.Title} | AdminTasks: {list.BaseTemplate}");
                        break;

                    default:
                        Console.WriteLine($"{list.Title} | Onbekend lijsttype: {list.BaseTemplate}");
                        break;
                }

            }
            context.ExecuteQuery();
        }

        private static void GetListProperties(ClientContext context, List list)
        {
            context.Load(
                            list,
                            l => l.AdditionalUXProperties,
                            l => l.AllowContentTypes,
                            l => l.AllowDeletion,
                            l => l.Author,
                            l => l.BaseTemplate,
                            l => l.BaseType,
                            l => l.BrowserFileHandling,
                            l => l.Color,
                            l => l.ContentTypes,
                            l => l.ContentTypesEnabled,
                            l => l.CrawlNonDefaultViews,
                            l => l.CreatablesInfo,
                            l => l.Created,
                            l => l.CurrentChangeToken,
                            l => l.CustomActionElements,
                            l => l.DataSource,
                            l => l.DefaultContentApprovalWorkflowId,
                            l => l.DefaultDisplayFormUrl,
                            l => l.DefaultEditFormUrl,
                            l => l.DefaultItemOpenInBrowser,
                            l => l.DefaultItemOpenUseListSetting,
                            l => l.DefaultNewFormUrl,
                            l => l.DefaultView,
                            l => l.DefaultViewPath,
                            l => l.DefaultViewUrl,
                            l => l.Description,
                            l => l.DescriptionResource,
                            l => l.Direction,
                            l => l.DisableCommenting,
                            l => l.DisableGridEditing,
                            l => l.DocumentTemplateUrl,
                            l => l.DraftVersionVisibility,
                            l => l.EffectiveBasePermissions,
                            l => l.EffectiveBasePermissionsForUI,
                            l => l.EnableAssignToEmail,
                            l => l.EnableAttachments,
                            l => l.EnableFolderCreation,
                            l => l.EnableMinorVersions,
                            l => l.EnableModeration,
                            l => l.EnableRequestSignOff,
                            l => l.EnableVersioning,
                            l => l.EntityTypeName,
                            l => l.EventReceivers,
                            l => l.ExcludeFromOfflineClient,
                            l => l.ExemptFromBlockDownloadOfNonViewableFiles,
                            l => l.Fields,
                            l => l.FileSavePostProcessingEnabled,
                            l => l.ForceCheckout,
                            l => l.Forms,
                            l => l.HasContentAssemblyTemplates,
                            l => l.HasCopyMoveRules,
                            l => l.HasExternalDataSource,
                            l => l.HasFolderColoringFields,
                            l => l.HasListBoundContentAssemblyTemplates,
                            l => l.Hidden,
                            l => l.HighPriorityMediaProcessing,
                            l => l.Icon,
                            l => l.Id,
                            l => l.ImagePath,
                            l => l.ImageUrl,
                            l => l.InformationRightsManagementSettings,
                            l => l.DefaultSensitivityLabelForLibrary,
                            l => l.SensitivityLabelToEncryptOnDOwnloadForLibrary,
                            l => l.IrmEnabled,
                            l => l.IrmExpire,
                            l => l.IrmReject,
                            l => l.IsApplicationList,
                            l => l.IsCatalog,
                            l => l.IsDefaultDocumentLibrary,
                            l => l.IsEnterpriseGalleryLibrary,
                            l => l.IsPredictionModelApplied,
                            l => l.IsPrivate,
                            l => l.IsSiteAssetsLibrary,
                            l => l.IsSystemList,
                            l => l.ItemCount,
                            l => l.LastItemDeletedDate,
                            l => l.LastItemModifiedDate,
                            l => l.LastItemUserModifiedDate,
                            l => l.ListExperienceOptions,
                            l => l.ListFormCustomized,
                            l => l.ListItemEntityTypeFullName,
                            l => l.ListSchemaVersion,
                            l => l.MajorVersionLimit,
                            l => l.MajorWithMinorVersionsLimit,
                            l => l.MultipleDataList,
                            l => l.NoCrawl,
                            l => l.OnQuickLaunch,
                            l => l.PageRenderType,
                            l => l.ParentWeb,
                            l => l.ParentWebPath,
                            l => l.ParentWebUrl,
                            l => l.ParserDisabled,
                            l => l.ReadSecurity,
                            l => l.RootFolder,
                            l => l.SchemaXml,
                            l => l.ServerTemplateCanCreateFolders,
                            l => l.ShowHiddenFieldsInModernForm,
                            l => l.TemplateFeatureId,
                            l => l.TemplateTypeId,
                            l => l.Title,
                            l => l.TitleResource,
                            l => l.UserCustomActions);
            context.ExecuteQuery();
        }
    }
}
