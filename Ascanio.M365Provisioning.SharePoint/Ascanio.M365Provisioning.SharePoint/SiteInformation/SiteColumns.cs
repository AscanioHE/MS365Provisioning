using Microsoft.SharePoint.Client;
using Ascanio.M365Provisioning.SharePoint.Services;

namespace Ascanio.M365Provisioning.SharePoint.SiteInformation
{
    public class SiteColumns
    {
        public SiteColumns() 
        { 
            SharePointService sharePointService = new();
            using(ClientContext context = sharePointService.GetClientContext())
            {
                Web web = context.Web;
                context.Load(
                            web,
                            w => w.Fields,
                            w => w.Fields.Include(f => f.Hidden,
                                                  f => f.InternalName,
                                                  f => f.SchemaXml,
                                                  f => f.DefaultValue
                                                  )
                            );
                context.ExecuteQuery();
                List<SiteColumnsDTO> siteColumnsDTO = new();

                foreach(Field siteColumn in web.Fields)
                {
                    if(!siteColumn.Hidden)
                    {
                        siteColumnsDTO.Add(new
                        (
                            siteColumn.Title,
                            siteColumn.SchemaXml,
                            siteColumn.DefaultValue
                        ));
                    }
                }

                WriteData2Json writeData2Json = new();
                writeData2Json.Write2JsonFile(siteColumnsDTO, sharePointService.SiteColumnsFilePath);
            }
        }

    }
}
