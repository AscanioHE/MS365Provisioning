using Microsoft.Extensions.Logging;

namespace MS365Provisioning.Common;

public interface IExportServices
{
    object DtoFile { get; set; }
    string FileName { get; set; }
    string JsonString { get; set; }

    string ConvertToJsonString();
    void ExportSettings(object dtoFile, string fileName, string jsonString);
    bool WriteJsonStringToFile();
}