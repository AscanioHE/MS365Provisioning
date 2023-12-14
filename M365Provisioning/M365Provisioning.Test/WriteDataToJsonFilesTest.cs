using M365Provisioning.SharePoint.Interfaces;
using M365Provisioning.SharePoint;
using WriteDataToJsonFiles;
using M365Provisioning.SharePoint.Functions;

namespace M365Provisioning.Test;

public class WriteDataToJsonFilesTest
{
    private SharePointServices sharePointServices = new();

    public WriteDataToJsonFilesTest()
    {
        WriteDataToJson = new WriteDataToJsonFile(sharePointServices.SiteSettingsFilePath);
    }

    private IWriteDataToJson WriteDataToJson { get; }

    [Fact]
    public void Try_Convert_DtoFile_To_String()
    {
        
    }
}