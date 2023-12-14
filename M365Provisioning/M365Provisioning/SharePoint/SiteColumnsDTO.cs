namespace M365Provisioning.SharePoint
{
    public class SiteColumnsDto
    {
        public string Name {  get; set; } 
        public string SchemaXml { get; set; }
        public string DefaultValue {  get; set; } 

        public SiteColumnsDto(string name, string schemaXml,string defaultValue) 
        { 
            Name = name;
            SchemaXml = schemaXml;
            DefaultValue = defaultValue;
        }

    }
}
