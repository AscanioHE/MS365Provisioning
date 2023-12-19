using Microsoft.Extensions.Logging;

namespace MS365Provisioning.Common;

public interface IExportSettings
{
    object DtoFile { get; set; }
    string FilePath { get; set; }
    string FileName { get; set; }

    string ConvertToJsonString();
    bool WriteJsonStringToFile();
}