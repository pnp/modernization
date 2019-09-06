using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using SharePointPnP.Modernization.Framework.Entities;
using SharePointPnP.Modernization.Framework.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SharePointPnP.Modernization.Framework.Pages
{
    public class WebPartPageOnPremises : WebPartPage
    {

        #region construction
        /// <summary>
        /// Instantiates a web part page object for on-premises environments
        /// </summary>
        /// <param name="page">ListItem holding the page to analyze</param>
        /// <param name="pageFile">File holding the page (for pages living outside of a library)</param>
        /// <param name="pageTransformation">Page transformation information</param>
        public WebPartPageOnPremises(ListItem page, File pageFile, PageTransformation pageTransformation) : base(page, pageFile, pageTransformation)
        {
        }
        #endregion


        /// <summary>
        /// Analyses a webpart page from on-premises environment
        /// </summary>
        /// <param name="includeTitleBarWebPart"></param>
        /// <returns></returns>
        public override Tuple<PageLayout, List<WebPartEntity>> Analyze(bool includeTitleBarWebPart = false)
        {
            List<WebPartEntity> webparts = new List<WebPartEntity>();

            //Load the page
            string webPartPageUrl = null;
            File webPartPage = null;

            if (this.page != null)
            {
                webPartPageUrl = page[Constants.FileRefField].ToString();
                webPartPage = cc.Web.GetFileByServerRelativeUrl(webPartPageUrl);
            }
            else
            {
                webPartPageUrl = this.pageFile.EnsureProperty(p => p.ServerRelativeUrl);
                webPartPage = this.pageFile;
            }

            // Load web parts on web part page
            // Note: Web parts placed outside of a web part zone using SPD are not picked up by the web part manager. There's no API that will return those,
            //       only possible option to add parsing of the raw page aspx file.
            var limitedWPManager = webPartPage.GetLimitedWebPartManager(PersonalizationScope.Shared);
            cc.Load(limitedWPManager);
            
            IEnumerable<WebPartDefinition> webParts = cc.LoadQuery(limitedWPManager.WebParts.IncludeWithDefaultProperties(wp => wp.Id, wp => wp.WebPart.Title, wp => wp.WebPart.ZoneIndex, wp => wp.WebPart.IsClosed, wp => wp.WebPart.Hidden));
            cc.ExecuteQueryRetry();

            List<WebServiceWebPartProperties> webServiceWebPartEntities = LoadWebPartPropertiesFromWebServices(webPartPage.EnsureProperty(p => p.ServerRelativeUrl)); ;
            var pageUrl = page[Constants.FileRefField].ToString();

            // Check page type
            var layout = GetLayoutFromWebServices(webPartPageUrl);

            if (webParts.Count() > 0)
            {
                List<WebPartPlaceHolder> webPartsToRetrieve = new List<WebPartPlaceHolder>();

                foreach (var foundWebPart in webParts)
                {
                    webPartsToRetrieve.Add(new WebPartPlaceHolder()
                    {
                        WebPartDefinition = foundWebPart,
                        WebPartXml = null,
                        WebPartType = "",
                    });
                }

                
                foreach (var foundWebPart in webPartsToRetrieve)
                {
                    var wsWp = webServiceWebPartEntities.FirstOrDefault(o => o.Id == foundWebPart.WebPartDefinition.Id);

                    // Skip Microsoft.SharePoint.WebPartPages.TitleBarWebPart webpart in TitleBar zone
                    if (wsWp.ZoneId.Equals("TitleBar", StringComparison.InvariantCultureIgnoreCase))
                    {
                        if (!includeTitleBarWebPart)
                        {
                            continue;
                        }
                    }

                    var wsExportMode = wsWp.Properties.FirstOrDefault(o => o.Key.Equals("exportmode", StringComparison.InvariantCultureIgnoreCase));

                    if (!string.IsNullOrEmpty(wsExportMode.Value) && wsExportMode.Value.Equals("all", StringComparison.InvariantCultureIgnoreCase))
                    {
                        var webPartXml = ExportWebPartXmlWorkaround(pageUrl, foundWebPart.WebPartDefinition.Id.ToString());
                        foundWebPart.WebPartXmlOnPremises = webPartXml;
                    }
                }
                

                foreach (var foundWebPart in webPartsToRetrieve)
                {
                    bool isExportable = false;
                    Dictionary<string, object> webPartProperties = null;

                    var wsWp = webServiceWebPartEntities.FirstOrDefault(o => o.Id == foundWebPart.WebPartDefinition.Id);
                    webPartProperties = wsWp.PropertiesAsStringObjectDictionary();

                    // Skip Microsoft.SharePoint.WebPartPages.TitleBarWebPart webpart in TitleBar zone
                    if (wsWp.ZoneId.Equals("TitleBar", StringComparison.InvariantCultureIgnoreCase))
                    {
                        if (!includeTitleBarWebPart)
                        {
                            continue;
                        }
                    }

                    // If the web service call includes the export mode value then set the export options
                    var wsExportMode = wsWp.Properties.FirstOrDefault(o => o.Key.Equals("exportmode", StringComparison.InvariantCultureIgnoreCase));

                    if (!string.IsNullOrEmpty(wsExportMode.Value) && wsExportMode.Value.ToString().Equals("all", StringComparison.InvariantCultureIgnoreCase))
                    {
                        isExportable = true;
                    }

                    if (!isExportable)
                    {
                        // Use different approach to determine type as we can't export the web part XML without indroducing a change
                        foundWebPart.WebPartType = GetTypeFromProperties(webPartProperties, true);
                    }
                    else
                    {
                        foundWebPart.WebPartType = GetType(foundWebPart.WebPartXmlOnPremises);
                    }

                    webparts.Add(new WebPartEntity()
                    {
                        Title = foundWebPart.WebPartDefinition.WebPart.Title,
                        Type = foundWebPart.WebPartType,
                        Id = foundWebPart.WebPartDefinition.Id,
                        ServerControlId = foundWebPart.WebPartDefinition.Id.ToString(),
                        Row = GetRow(wsWp.ZoneId, layout),
                        Column = GetColumn(wsWp.ZoneId, layout),
                        Order = foundWebPart.WebPartDefinition.WebPart.ZoneIndex,
                        ZoneId = wsWp.ZoneId,
                        ZoneIndex = (uint)foundWebPart.WebPartDefinition.WebPart.ZoneIndex,
                        IsClosed = foundWebPart.WebPartDefinition.WebPart.IsClosed,
                        Hidden = foundWebPart.WebPartDefinition.WebPart.Hidden,
                        Properties = Properties(webPartProperties, foundWebPart.WebPartType, foundWebPart.WebPartXmlOnPremises),
                    });
                }
            }

            return new Tuple<PageLayout, List<WebPartEntity>>(layout, webparts);
        }

        /// <summary>
        /// Gets and parses the layout from the web services URL
        /// </summary>
        /// <param name="webPartPageUrl"></param>
        /// <returns></returns>
        internal PageLayout GetLayoutFromWebServices(string webPartPageUrl)
        {
            var wsPageDocument = ExtractWebPartDocumentViaWebServicesFromPage(webPartPageUrl);

            if (!string.IsNullOrEmpty(wsPageDocument.Item1))
            {
                //Example fragment from WS
                //<li>vti_setuppath
                //<li>SR|1033&#92;STS&#92;doctemp&#92;smartpgs&#92;spstd2.aspx
                //<li>vti_generator

                var fullDocument = wsPageDocument.Item1;

                if (!string.IsNullOrEmpty(fullDocument))
                {
                    if (fullDocument.ContainsIgnoringCasing(@"1033&#92;STS&#92;doctemp&#92;smartpgs&#92;spstd1.aspx"))
                    {
                        return PageLayout.WebPart_FullPageVertical;
                    }
                    else if (fullDocument.ContainsIgnoringCasing(@"1033&#92;STS&#92;doctemp&#92;smartpgs&#92;spstd2.aspx"))
                    {
                        return PageLayout.WebPart_HeaderFooterThreeColumns;
                    }
                    else if (fullDocument.ContainsIgnoringCasing(@"1033&#92;STS&#92;doctemp&#92;smartpgs&#92;spstd3.aspx"))
                    {
                        return PageLayout.WebPart_HeaderLeftColumnBody;
                    }
                    else if (fullDocument.ContainsIgnoringCasing(@"1033&#92;STS&#92;doctemp&#92;smartpgs&#92;spstd4.aspx"))
                    {
                        return PageLayout.WebPart_HeaderRightColumnBody;
                    }
                    else if (fullDocument.ContainsIgnoringCasing(@"1033&#92;STS&#92;doctemp&#92;smartpgs&#92;spstd5.aspx"))
                    {
                        return PageLayout.WebPart_HeaderFooter2Columns4Rows;
                    }
                    else if (fullDocument.ContainsIgnoringCasing(@"1033&#92;STS&#92;doctemp&#92;smartpgs&#92;spstd6.aspx"))
                    {
                        return PageLayout.WebPart_HeaderFooter4ColumnsTopRow;
                    }
                    else if (fullDocument.ContainsIgnoringCasing(@"1033&#92;STS&#92;doctemp&#92;smartpgs&#92;spstd7.aspx"))
                    {
                        return PageLayout.WebPart_LeftColumnHeaderFooterTopRow3Columns;
                    }
                    else if (fullDocument.ContainsIgnoringCasing(@"1033&#92;STS&#92;doctemp&#92;smartpgs&#92;spstd8.aspx"))
                    {
                        return PageLayout.WebPart_RightColumnHeaderFooterTopRow3Columns;
                    }
                    else if (fullDocument.ContainsIgnoringCasing(@"SiteTemplates&#92;STS&#92;default.aspx"))
                    {
                        return PageLayout.WebPart_2010_TwoColumnsLeft;
                    }
                }

            }

            return PageLayout.WebPart_Custom;
        }
    }
}
