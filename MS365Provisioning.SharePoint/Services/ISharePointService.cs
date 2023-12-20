using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MS365Provisioning.SharePoint.Model;

namespace MS365Provisioning.SharePoint.Services
{
    public interface ISharePointService
    {
        List<SiteSettingsDto> LoadSiteSettings();
        List<ListsSettingsDto> LoadListsSettings();
        List<ListViewDto> LoadListViews();
        List<SiteColumnsDto> LoadSiteColumns();
    }
}
