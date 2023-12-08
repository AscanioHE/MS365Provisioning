namespace Ascanio.M365Provisioning.SharePoint.SiteInformation
{
    public interface ISiteSettingsService
    {
        List<SiteSettingsDTO> Load();
        void WriteToJsonFile(List<SiteSettingsDTO> webTemplatesDTO, string jsonFilePath);

        string ConvertToJson(List<SiteSettingsDTO> webTemplatesDTO);
    }
}
