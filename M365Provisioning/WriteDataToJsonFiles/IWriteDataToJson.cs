using WriteDataToJsonFiles;

namespace WriteDataToJsonFiles;

public interface IWriteDataToJson
{
    string ConvertDtoToString();
    string Write2JsonFile();
}