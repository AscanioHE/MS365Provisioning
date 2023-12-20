namespace MS365Provisioning.SharePoint.Model;

public class ContentTypesDto
{
    public string Title { get; set; }
    public string ParentCt { get; set; }
    public List<string> FieldTitle { get; set; }
    public bool Required { get; set; }

    public ContentTypesDto(string columnName, string parentCt, List<string> fieldTitle, bool required)
    {
        Title = columnName;
        ParentCt = parentCt;
        FieldTitle = fieldTitle;
        Required = required;
    }
    
}