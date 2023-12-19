namespace MS365Provisioning.SharePoint.Model;

public class SiteColumnsDto
{
    public string ColumnName { get; set; }
    public string SchemaXml { get; set; }
    public string DefaultValue { get; set; }

    public SiteColumnsDto(string columnName, string schemaXml, string defaultValue)
    {
        ColumnName = columnName;
        SchemaXml = schemaXml;
        DefaultValue = defaultValue;
    }

}