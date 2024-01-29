using MS365Provisioning.SharePoint.Model;

namespace MS365Provisioning.Common.Tests
{
    public class ExportServicesTest
    {
        public object DtoFile { get; set; }
        public string FileName { get; set; }
        public string JsonString { get; set; }

        private readonly IExportServices _exportSettings;

        public ExportServicesTest()
        {
            List<ListViewDto> dtoTestFile = new();
            DtoFile = dtoTestFile;
            _exportSettings = new ExportServices();
            FileName = _exportSettings.FileName;
            JsonString = _exportSettings.ConvertToJsonString();
        }

        [Fact]
        public void Try_ConvertToJsonString_Expect_String()
        {
            //Act
            JsonString = _exportSettings.ConvertToJsonString();
            //Assert
            Assert.IsType<string>(JsonString);
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