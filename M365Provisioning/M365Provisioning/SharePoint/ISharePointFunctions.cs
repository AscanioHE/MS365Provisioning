﻿using M365Provisioning.SharePoint;
using Microsoft.SharePoint.Client;
using PnP.Framework.Provisioning.Model;
using System.Collections.Generic;
using MS365Provisioning.SharePoint.Model;

namespace M365Provisioning.SharePoint;

public interface ISharePointFunctions
{
   List<SiteSettingsDto> LoadSiteSettings();
   List<ListsSettingsDto> LoadListsSettings();
   List<ListViewDto> LoadListViews();
   List<SiteColumnsDto> LoadSiteColumnsDtos();
}