using Microsoft.Extensions.Logging;
using MS365Provisioning.Common;
using Newtonsoft.Json;

namespace MS365Provisioning.Common
{
    public class ExportServices : IExportServices
    {
        public object DtoFile { get; set; } = new object();
        public string FileName { get; set; } = string.Empty;
        public string JsonString { get; set; } = string.Empty;

        public string ConvertToJsonString()
        {
            JsonString = JsonConvert.SerializeObject(DtoFile, Formatting.Indented);
            return JsonString;
        }

        public bool WriteJsonStringToFile()
        {
            try
            {
                File.WriteAllText(FileName!, JsonString + Environment.NewLine);
                return true;
            }
            catch
            {
                return false;
            }
        }

        void IExportServices.ExportSettings(object dtoFile, string fileName, string jsonString)
        {
            DtoFile = dtoFile;
            FileName = fileName;
            JsonString = jsonString;
        }
    }
}