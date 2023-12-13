using Newtonsoft.Json;
using File = System.IO.File;

namespace WriteDataToJsonFiles
{
    public class WriteDataToJsonFile : IWriteDataToJson
    {
        public WriteDataToJsonFile()
        {
        }

        public string ConvertDtoToString(object dtoFile)
        {
            string json = JsonConvert.SerializeObject(dtoFile, Formatting.Indented);
            return json;
        }

        public string ConvertDtoToString()
        {
            throw new NotImplementedException();
        }
    }
}