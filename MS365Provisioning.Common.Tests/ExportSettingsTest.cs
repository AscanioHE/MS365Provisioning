using Microsoft.Extensions.Logging;
using MS365Provisioning.SharePoint.Model;

namespace MS365Provisioning.Common.Tests
{
    public class ExportSettingsTest : IExportSettings
    {
        public object DtoFile { get; set; }
        public string FileName {  get; set; }
        public ILogger _logger;

        public ExportSettingsTest(object dtoFile, string fileName,ILogger logger)
        {
            _logger = logger;
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