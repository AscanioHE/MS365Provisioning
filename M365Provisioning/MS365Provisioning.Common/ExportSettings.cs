using Microsoft.Extensions.Logging;
using MS365Provisioning.Common;

namespace MS365Provisioning.Common
{
    public class ExportSettings : IExportSettings
    {
        public string JsonString { get; set; }
        public ILogger _logger;

        public ExportSettings(string json)
        {
            JsonString = json;
        }

        public string ConvertToJsonString()
        {
            throw new NotImplementedException();
        }

        public bool WriteJsonStringToFile()
        {
            throw new NotImplementedException();
        }
    }
}