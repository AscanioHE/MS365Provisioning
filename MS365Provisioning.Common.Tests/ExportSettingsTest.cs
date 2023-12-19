using MS365Provisioning.SharePoint.Model;

namespace MS365Provisioning.Common.Tests
{
    public class ExportSettingsTest : IExportSettings
    {
        public object DtoFile { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public string FilePath { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public string FileName { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

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