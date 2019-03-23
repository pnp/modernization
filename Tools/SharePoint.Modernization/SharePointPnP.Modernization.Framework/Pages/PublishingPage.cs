using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using SharePointPnP.Modernization.Framework.Entities;
using SharePointPnP.Modernization.Framework.Functions;
using SharePointPnP.Modernization.Framework.Publishing;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SharePointPnP.Modernization.Framework.Pages
{
    /// <summary>
    /// Analyzes a publishing page
    /// </summary>
    public class PublishingPage : BasePage
    {
        private PublishingPageTransformation publishingPageTransformation;
        private PublishingFunctionProcessor functionProcessor;

        #region Construction
        /// <summary>
        /// Instantiates a publishing page object
        /// </summary>
        /// <param name="page">ListItem holding the page to analyze</param>
        /// <param name="pageTransformation">Page transformation information</param>
        public PublishingPage(ListItem page, PageTransformation pageTransformation) : base(page, pageTransformation)
        {
            // no PublishingPageTransformation specified, fall back to default
            this.publishingPageTransformation = new PageLayoutManager(cc).LoadDefaultPageLayoutMappingFile();
            this.functionProcessor = new PublishingFunctionProcessor(page, cc, this.publishingPageTransformation);
        }

        /// <summary>
        /// Instantiates a publishing page object
        /// </summary>
        /// <param name="page">ListItem holding the page to analyze</param>
        /// <param name="pageTransformation">Page transformation information</param>
        public PublishingPage(ListItem page, PageTransformation pageTransformation, PublishingPageTransformation publishingPageTransformation) : base(page, pageTransformation)
        {
            this.publishingPageTransformation = publishingPageTransformation;
            this.functionProcessor = new PublishingFunctionProcessor(page, cc, this.publishingPageTransformation);
        }
        #endregion

        /// <summary>
        /// Analyses a publishing page
        /// </summary>
        /// <returns>Information about the analyzed publishing page</returns>
        public Tuple<PageLayout, List<WebPartEntity>> Analyze()
        {
            List<WebPartEntity> webparts = new List<WebPartEntity>();
            PageLayout layout = PageLayout.Wiki_OneColumn;

            //Load the page
            var publishingPageUrl = page[Constants.FileRefField].ToString();
            var publishingPage = cc.Web.GetFileByServerRelativeUrl(publishingPageUrl);

            // Load page properties
            //var pageProperties = publishingPage.Properties;
            //cc.Load(pageProperties);

            // Load relevant model data for the used page layout
            string usedPageLayout = System.IO.Path.GetFileNameWithoutExtension(page.PageLayoutFile());
            var publishingPageTransformationModel = this.publishingPageTransformation.PageLayouts.Where(p => p.Name.Equals(usedPageLayout, StringComparison.InvariantCultureIgnoreCase)).First();

            if (publishingPageTransformationModel == null)
            {
                throw new Exception($"No valid page transformation model could be retrieved for publishing page layout {usedPageLayout}");
            }

            #region Process fields that become web parts 
            if (publishingPageTransformationModel.WebParts != null)
            {
                #region Publishing Html column processing
                // Converting to WikiTextPart is a special case as we'll need to process the html
                var wikiTextWebParts = publishingPageTransformationModel.WebParts.Where(p => p.TargetWebPart.Equals("SharePointPnP.Modernization.WikiTextPart", StringComparison.InvariantCultureIgnoreCase));
                List<WebPartPlaceHolder> webPartsToRetrieve = new List<WebPartPlaceHolder>();
                foreach (var wikiTextPart in wikiTextWebParts)
                {
                    var pageContents = page.FieldValues[wikiTextPart.Name].ToString();
                    var htmlDoc = parser.Parse(pageContents);


                    // Analyze the html block (which is a wiki block)
                    var content = htmlDoc.FirstElementChild.LastElementChild;
                    AnalyzeWikiContentBlock(webparts, htmlDoc, webPartsToRetrieve, 1, 1, content);
                }

                // Bulk load the needed web part information
                if (webPartsToRetrieve.Count > 0)
                {
                    LoadWebPartsInWikiContentFromServer(webparts, publishingPage, webPartsToRetrieve);
                }
                #endregion

                #region Generic processing of the other 'webpart' fields
                var fieldWebParts = publishingPageTransformationModel.WebParts.Where(p => !p.TargetWebPart.Equals("SharePointPnP.Modernization.WikiTextPart", StringComparison.InvariantCultureIgnoreCase));
                foreach(var fieldWebPart in fieldWebParts)
                {
                    Dictionary<string, string> properties = new Dictionary<string, string>();

                    foreach(var fieldWebPartProperty in fieldWebPart.Property)
                    {
                        if (!string.IsNullOrEmpty(fieldWebPartProperty.Functions))
                        {
                            // execute function
                            var evaluatedField = this.functionProcessor.Process(fieldWebPartProperty);
                            if (!string.IsNullOrEmpty(evaluatedField.Item1) && !properties.ContainsKey(evaluatedField.Item1))
                            {
                                properties.Add(evaluatedField.Item1, evaluatedField.Item2);
                            }
                        }
                        else
                        {
                            properties.Add(fieldWebPartProperty.Name, page.FieldValues[fieldWebPart.Name].ToString().Trim());
                        }
                    }

                    var wpEntity = new WebPartEntity()
                    {
                        Title = fieldWebPart.Name,
                        Type = fieldWebPart.TargetWebPart,
                        Id = Guid.Empty,
                        Row = 1,
                        Column = 1,
                        Order = 1,
                        Properties = properties,
                    };

                    webparts.Add(wpEntity);

                }
            }
            #endregion
            #endregion

            #region Web Parts in webpart zone handling
            // Load web parts put in web part zones on the publishing page
            // Note: Web parts placed outside of a web part zone using SPD are not picked up by the web part manager. 
            var limitedWPManager = publishingPage.GetLimitedWebPartManager(PersonalizationScope.Shared);
            cc.Load(limitedWPManager);

            IEnumerable<WebPartDefinition> webPartsViaManager = cc.LoadQuery(limitedWPManager.WebParts.IncludeWithDefaultProperties(wp => wp.Id, wp => wp.ZoneId, wp => wp.WebPart.ExportMode, wp => wp.WebPart.Title, wp => wp.WebPart.ZoneIndex, wp => wp.WebPart.IsClosed, wp => wp.WebPart.Hidden, wp => wp.WebPart.Properties));
            cc.ExecuteQueryRetry();

            if (webPartsViaManager.Count() > 0)
            {
                List<WebPartPlaceHolder> webPartsToRetrieve = new List<WebPartPlaceHolder>();

                foreach (var foundWebPart in webPartsViaManager)
                {
                    // Remove the web parts which we've already picked up by analyzing the wiki content block
                    if (webparts.Where(p => p.Id.Equals(foundWebPart.Id)).First() != null)
                    {
                        continue;
                    }

                    webPartsToRetrieve.Add(new WebPartPlaceHolder()
                    {
                        WebPartDefinition = foundWebPart,
                        WebPartXml = null,
                        WebPartType = "",
                    });
                }

                bool isDirty = false;
                foreach (var foundWebPart in webPartsToRetrieve)
                {
                    if (foundWebPart.WebPartDefinition.WebPart.ExportMode == WebPartExportMode.All)
                    {
                        foundWebPart.WebPartXml = limitedWPManager.ExportWebPart(foundWebPart.WebPartDefinition.Id);
                        isDirty = true;
                    }
                }
                if (isDirty)
                {
                    cc.ExecuteQueryRetry();
                }

                foreach (var foundWebPart in webPartsToRetrieve)
                {
                    if (foundWebPart.WebPartDefinition.WebPart.ExportMode != WebPartExportMode.All)
                    {
                        // Use different approach to determine type as we can't export the web part XML without indroducing a change
                        foundWebPart.WebPartType = GetTypeFromProperties(foundWebPart.WebPartDefinition.WebPart.Properties);
                    }
                    else
                    {
                        foundWebPart.WebPartType = GetType(foundWebPart.WebPartXml.Value);
                    }

                    webparts.Add(new WebPartEntity()
                    {
                        Title = foundWebPart.WebPartDefinition.WebPart.Title,
                        Type = foundWebPart.WebPartType,
                        Id = foundWebPart.WebPartDefinition.Id,
                        ServerControlId = foundWebPart.WebPartDefinition.Id.ToString(),
                        Row = 1,
                        Column = 1,
                        Order = foundWebPart.WebPartDefinition.WebPart.ZoneIndex,
                        ZoneId = foundWebPart.WebPartDefinition.ZoneId,
                        ZoneIndex = (uint)foundWebPart.WebPartDefinition.WebPart.ZoneIndex,
                        IsClosed = foundWebPart.WebPartDefinition.WebPart.IsClosed,
                        Hidden = foundWebPart.WebPartDefinition.WebPart.Hidden,
                        Properties = Properties(foundWebPart.WebPartDefinition.WebPart.Properties, foundWebPart.WebPartType, foundWebPart.WebPartXml == null ? "" : foundWebPart.WebPartXml.Value),
                    });
                }
            }
            #endregion

            return new Tuple<PageLayout, List<WebPartEntity>>(layout, webparts);
        }
    }
}
