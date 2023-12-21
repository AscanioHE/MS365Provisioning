using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using MS365Provisioning.SharePoint.Model;
using MS365Provisioning.SharePoint.Services;

namespace MS365Provisioning.Common.Tests
{
    public class ExportSettingsTest
    {
        public object DtoFile { get; set; }
        public string FileName {  get; set; }
        public ILogger _logger;

        private readonly IExportSettings _exportSettings;

        public ExportSettingsTest(object dtoFile, string fileName,ILogger logger)
        {
            //Arrange
            _logger = logger;
            FileName = fileName;
            DtoFile = dtoFile;
            _exportSettings = new ExportSettings(this, FileName, _logger);
        }

        [Fact]
        public void Try_ConvertToJsonString_Expect_String()
        {
            //Act
            string json = _exportSettings.ConvertToJsonString();
            //Assert
            Assert.IsType<string>(json);
        }
        [Fact]
        public void Try_WriteJsonStringToFile_Expect_Bool()
        {
            //Act
            bool success = _exportSettings.WriteJsonStringToFile();
            //Assert
            Assert.IsType<bool>(success);
        }
    }
}