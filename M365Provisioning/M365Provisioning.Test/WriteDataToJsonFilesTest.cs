using M365Provisioning.SharePoint.Interfaces;
using M365Provisioning.SharePoint;
using WriteDataToJsonFiles;
using M365Provisioning.SharePoint.Functions;
using Newtonsoft.Json;

namespace M365Provisioning.Test;

public class WriteDataToJsonFilesTest
{
    private SharePointServices sharePointServices = new();

    [Fact]
    public void Try_Convert_DtoFile_Expect_String()
    {
        //Arrange
        string tempFilePath = Path.GetTempFileName();
        object validDto = new();
        //Act
        WriteDataToJsonFile writeDataToJson = new()
        {
            DtoFile = validDto,
            JsonFilePath = tempFilePath
        };
        string fileContents = File.ReadAllText(tempFilePath);
        //Assert
        string expectedJsonString = JsonConvert.SerializeObject(validDto, Formatting.Indented);
        Assert.Equal(expectedJsonString, fileContents);

        File.Delete(tempFilePath);
    }
}