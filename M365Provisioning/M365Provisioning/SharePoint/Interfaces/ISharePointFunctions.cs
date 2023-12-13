using M365Provisioning.SharePoint.DTO;

namespace M365Provisioning.SharePoint.Interfaces;

public interface ISharePointFunctions
{
    List<SiteSettingsDto> LoadSiteSettings();
    List<ListDto> GetLists();

}