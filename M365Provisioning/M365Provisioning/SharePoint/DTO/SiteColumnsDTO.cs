namespace Ascanio.M365Provisioning.SharePoint.SiteInformation
{
    public class SiteColumnsDTO
    {
        public string Name {  get; set; }  = string.Empty;
        public string SchemaXML { get; set; } = string.Empty;  
        public string DefaultValue {  get; set; } = string.Empty;

        public SiteColumnsDTO(string name, string schemaXML,string defaultValue) 
        { 
            Name = name;
            SchemaXML = schemaXML;
            DefaultValue = defaultValue;
        }

    }
}
