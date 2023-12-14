using M365Provisioning.SharePoint;
using PnP.Framework.Provisioning.Model;
using System.Collections.Generic;

namespace M365Provisioning.SharePoint;

public interface ISharePointFunctions
{
   List<SiteSettingsDto> LoadSiteSettings();
   List<ListsSettingsDto> LoadListsSettings();

}