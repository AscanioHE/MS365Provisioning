using M365Provisioning.SharePoint.DTO;
using M365Provisioning.SharePoint.Services;

namespace M365Provisioning;

internal static class Program
{
    public static void Main()
    {
        try
        {
            _ = new SharePointServices().GetSiteSettings() ;
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
    }
}