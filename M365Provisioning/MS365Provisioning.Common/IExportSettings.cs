namespace MS365Provisioning.Common;

public interface IExportSettings
{
    string ConvertToJsonString();
    bool WriteJsonStringToFile();
}