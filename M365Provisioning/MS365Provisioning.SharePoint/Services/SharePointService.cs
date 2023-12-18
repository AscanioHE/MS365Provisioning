﻿using System.Collections;
using System.Diagnostics;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using MS365Provisioning.SharePoint.Model;
using MS365Provisioning.SharePoint.Settings;
using PnP.Core.Model.SharePoint;

namespace MS365Provisioning.SharePoint.Services
{
    public class SharePointService : ISharePointService
    {
        private readonly ISharePointSettingsService? _sharePointSettingsService;
        private readonly ILogger? _logger;
        private readonly ClientContext _clientContext;
        private readonly ListCollection? _lists;
        public SharePointService(ISharePointSettingsService? sharePointSettingsService, ILogger? logger, string siteUrl)
        {
            _sharePointSettingsService = sharePointSettingsService;
            _logger = logger;
            _clientContext = GetClientContext(siteUrl)!; 
            _lists = _clientContext.Web.Lists;
        }
        public SharePointService(ClientContext clientContext)
        {
            _clientContext = clientContext;
        }
        /*______________________________________________________________________________________________________________
         Create ClientContext
        ________________________________________________________________________________________________________________*/
        private ClientContext? GetClientContext(string siteUrl)
        {
            string message = $"{nameof(GetClientContext)} for site {siteUrl}...";
            _logger?.LogInformation(message: message);
            if (_sharePointSettingsService != null)
            {
                SharePointSettings? sharePointSettings = _sharePointSettingsService.GetSharePointSettings();

                ClientContext? ctx;
                using (X509Certificate2? certificate = GetCertificateByThumbprint(sharePointSettings.ThumbPrint))
                {
                    ctx = null;
                    try
                    {
                        PnP.Framework.AuthenticationManager authManager = new(sharePointSettings.ClientId, certificate,
                            sharePointSettings.TenantId);
                        ctx = authManager.GetContext(sharePointSettings.SiteUrl);
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogError(message: $"Error fetching the ClientContext : {ex.Message}");
                        return new ClientContext("");
                    }
                }
                return ctx;
            }
            else
            {
                return new ClientContext("");
            }
        }

        /*______________________________________________________________________________________________________________
         Fetch SiteSettings
        ________________________________________________________________________________________________________________*/

        public List<SiteSettingsDto> LoadSiteSettings()
        {
            List<SiteSettingsDto> siteSettingsDto = new();
            try
            {
                WebTemplateCollection webTemplateCollection = _clientContext.Web.GetAvailableWebTemplates(1033, true);
                _clientContext.Load(webTemplateCollection);
                _clientContext.ExecuteQuery();
                foreach (WebTemplate webTemplate in webTemplateCollection)
                {
                    siteSettingsDto.Add(new SiteSettingsDto
                    {
                        SiteTemplate = webTemplate.Name,
                        Value = webTemplate.Lcid
                    });
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError(message: $"Error fetching the Webtemplates : {ex.Message}");
            }
            finally
            {
                _clientContext.Dispose();
            }
            return siteSettingsDto;
        }

        /*______________________________________________________________________________________________________________
         Fetch Lists Settings
        ________________________________________________________________________________________________________________*/
        public List<ListsSettingsDto> LoadListsSettings()
        {
            List<ListsSettingsDto> listsSettingsDto = new();
            bool breakRoleAssignment = false;
            _clientContext.Load(_lists, lc => lc.Include(
                l => l.Hidden)
            );
            try
            {
                _clientContext.ExecuteQuery();
                if (_lists == null || _lists.Count <= 0) return listsSettingsDto;
                foreach (List list in _lists)
                {
                    if (!list.Hidden)
                    {
                        _clientContext.Load(
                            list,
                            l=>l.Title,
                            l => l.DefaultViewUrl,
                            l => l.BaseType,
                            l => l.ContentTypes,
                            l => l.OnQuickLaunch,
                            l => l.HasUniqueRoleAssignments,
                            l=> l.EnableFolderCreation,
                            l=>l.RoleAssignments,
                            l => l.Fields.Include(
                                f=> f.InternalName,
                                f=> f.Title)); 
                        try
                        {
                            _clientContext.ExecuteQuery();
                            List<string> contentTypes = GetListContentTypes(list);
                            Dictionary<string, string> listPermissions = GetPermissionDetails(list);
                            Guid enterpiseKeywordsValue = GetEnterpriseKeywordsValue();
                            List<string> quickLaunchHeaders = GetQuickLaunchHeaders();
                            RoleAssignmentCollection roleAssignmentCollection = list.RoleAssignments;
                            foreach (RoleAssignment roleAssignment in roleAssignmentCollection)
                            {
                                breakRoleAssignment = roleAssignment.RoleDefinitionBindings.AreItemsAvailable;
                            }
                            try
                            {
                                listsSettingsDto.Add(new
                                (
                                    list.Title,
                                    list.DefaultViewUrl,
                                    list.BaseType.ToString(),
                                    contentTypes,
                                    list.OnQuickLaunch,
                                    quickLaunchHeaders,
                                    list.EnableFolderCreation,
                                    enterpiseKeywordsValue,
                                    breakRoleAssignment,
                                    listPermissions
                                ));
                            }
                            catch (Exception ex)
                            {
                                _logger?.LogInformation(
                                    $"Unable to create the List Data Transfer Object : {ex.Message}");
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger?.LogInformation(message: $"Error Fetching list properties : {ex.Message}");
                        }
                    }
                }
                return listsSettingsDto;
            }
            catch (Exception ex)
            {
                _logger?.LogError($"Error fetching the ClientContext Lists: {ex.Message}");
                return new List<ListsSettingsDto>();
            }
        }
        
        private List<string> GetQuickLaunchHeaders()
        {
            List<string> quickLaunchHeaders = new();
            try
            {
                _clientContext.Load(_clientContext.Web.Navigation.QuickLaunch);
                _clientContext.ExecuteQuery();
                foreach (NavigationNode navigationNode in _clientContext.Web.Navigation.QuickLaunch)
                {
                    _clientContext.Load
                    (
                        navigationNode,
                        n => n.Children
                    );
                    try
                    {
                        _clientContext.ExecuteQuery();
                        foreach (NavigationNode childNode in navigationNode.Children)
                        {
                            quickLaunchHeaders.Add(childNode.Title);
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogInformation($"Error fetching Web Navigation QuickLaunch: {ex}");
                        
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.LogInformation($"Error fetching List QuickLaunchHeader : {ex.Message}");
            }

            return quickLaunchHeaders;
        }

        private Guid GetEnterpriseKeywordsValue()
        {
            Guid enterpriseKeywordsValue = Guid.Empty;
            try
            {
                Field enterpriseKeywords = _clientContext.Web.Fields.GetByInternalNameOrTitle("EnterpriseKeywords");
                if (enterpriseKeywords != null)
                {
                    _clientContext.Load(enterpriseKeywords);
                    _clientContext.ExecuteQuery();
                    enterpriseKeywordsValue = enterpriseKeywords.Id;
                }
            }
            catch (Exception ex)
            {
                // Log the exception
                _logger?.LogInformation(message: $"Error fetching Enterprise Keywords value: {ex.Message}");
                enterpriseKeywordsValue = Guid.Empty;
            }
            return enterpriseKeywordsValue;
        }
        Dictionary<string, string> GetPermissionDetails(List list)
        {
            IQueryable<RoleAssignment> queryForList = list.RoleAssignments.Include(
                roleAsg => roleAsg.Member,
                roleAsg => roleAsg.RoleDefinitionBindings.Include(roleDef => roleDef.Name));
            IEnumerable roles = _clientContext.LoadQuery(queryForList);
            try
            {
                _clientContext.ExecuteQuery();

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
                _logger?.LogInformation(message: $"Error fetching permissions : {ex}");
                return new Dictionary<string, string>();
            }
        }
        /*______________________________________________________________________________________________________________
         Fetch Lists List Views
        ________________________________________________________________________________________________________________*/
        public List<ListViewDto> LoadListViews()
        {
            throw new NotImplementedException();
        }
        public List<SiteColumnsDto> LoadSiteColumnsDtos()
        {
            throw new NotImplementedException();
        }
        private X509Certificate2? GetCertificateByThumbprint(string? thumbprint)
        {
            try
            {
                using X509Store store = new(StoreName.My, StoreLocation.CurrentUser);
                store.Open(OpenFlags.ReadOnly);
                Debug.Assert(thumbprint != null, nameof(thumbprint) + " != null");
                var certificates = store.Certificates.Find(X509FindType.FindByThumbprint, thumbprint, false);
                if (certificates.Count > 0)
                {
                    _logger?.LogInformation("Authenticated and connected to SharePoint!");
                    return certificates[0];
                }
                else
                {
                    throw new InvalidOperationException($"Certificate with thumbprint {thumbprint} not found!");
                }
            }
            catch (Exception ex)
            {
                _logger?.LogInformation($"Error creating a Certificate : {ex}");
                return null;
            }
        }
        private List<string> GetListContentTypes(List list)
        {
            List<string> contentTypes = new();
            try
            {
                _clientContext.Load(list.ContentTypes);
                _clientContext.ExecuteQuery();
                if (list.ContentTypes.Count == 0)
                {
                    return contentTypes;
                }
                foreach (ContentType contentType in list.ContentTypes)
                {
                    contentTypes.Add(contentType.Name);
                }
            }
            catch (Exception ex)
            {
                _logger?.LogInformation($"Error fetching ContentTypes: {ex.Message}");
                contentTypes.Clear();
            }
            return contentTypes;
        }
    }

    public class ExportFunctions
    {

    }
}
