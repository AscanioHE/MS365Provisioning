using MS365Provisioning.SharePoint.Model;

namespace MS365Provisioning.SharePoint.Services
{
    public interface ISharePointService
    {
        List<ListsSettingsDto> GetListsSettings();
    }
}