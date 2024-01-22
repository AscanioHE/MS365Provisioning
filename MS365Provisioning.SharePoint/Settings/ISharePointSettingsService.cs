using MS365Provisioning.SharePoint.Model;

namespace MS365Provisioning.SharePoint.Settings
{
    public interface ISharePointSettingsService
    {
        SharePointSettings GetSharePointSettings();
        FileSettings GetFileSettings();
    }
}
