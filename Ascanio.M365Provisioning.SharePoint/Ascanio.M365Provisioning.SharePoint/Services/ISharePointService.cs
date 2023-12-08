using Ascanio.M365Provisioning.SharePoint.SiteInformation;

namespace Ascanio.M365Provisioning.SharePoint.Services
{
    public interface ISharePointService
    {
        List<SiteSettingsDTO> GetSiteSettings();
        List<ListDTO> GetLists();
        List<ListViewDTO> GetListViews();

        string ConvertToJson(object o);
    }
}
