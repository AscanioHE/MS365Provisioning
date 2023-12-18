using System.Runtime.CompilerServices;
using M365Provisioning.SharePoint;
using WriteDataToJsonFiles;
using Newtonsoft.Json;
using System.IO;
using Microsoft.Extensions.Logging;
using Xunit;

namespace M365Provisioning.Test;

public class WriteDataToJsonFilesTest
{
    //private readonly IWriteDataToJson _writeDataToJson = new WriteDataToJsonFile(new ILogger());

    [Fact]
    public void Try_Convert_DtoFile_Expect_String()
    {
        //Arrange
        //string json = _writeDataToJson.ConvertDtoToString();
        //Assert
        //Assert.IsType<string>(json);
    }

    [Fact]
    public void Try_Write_Dto_File_Expect_String()
    {
        //Arrange
        //string json = _writeDataToJson.Write2JsonFile();
        //Assert
        //Assert.IsType<string>(json);
    }
}
