namespace MS365Provisioning.Common.Tests;

public interface IExportSharePointSettings
{
    string ConvertToJsonString();
    bool WriteJsonStringToFile();
}