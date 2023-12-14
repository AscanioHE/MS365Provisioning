namespace M365Provisioning.SharePoint;

public class ContentTypesDto
{
    public string Title { get; set; }
    public string ParentCt { get; set; }
    public string FieldTitle { get; set; }
    public bool Required { get; set; }

    public ContentTypesDto(string columnName, string parentCt, string fieldTitle, bool required)
    {
        Title = columnName;
        ParentCt = parentCt;
        FieldTitle = fieldTitle;
        Required = required;
    }
    
}