using Microsoft.SharePoint.Client;
using SharePointPnP.Modernization.Framework.Entities;
using SharePointPnP.Modernization.Framework.Telemetry;
using SharePointPnP.Modernization.Framework.Transform;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using AngleSharp;
using AngleSharp.Dom;
using AngleSharp.Dom.Html;
using AngleSharp.Parser.Html;

using ContentType = Microsoft.SharePoint.Client.ContentType;
using File = Microsoft.SharePoint.Client.File;
using SharePointPnP.Modernization.Framework.Publishing.Layouts;
using System.Text.RegularExpressions;

namespace SharePointPnP.Modernization.Framework.Publishing
{
    public class PageLayoutAnalyser : BaseTransform
    {
        /*
         * Plan
         *  Read a publishing page or read all the publishing page layouts - need to consider both options
         *  Validate that the client context is a publishing site
         *  Determine page layouts and the associated content type
         *  - Using web part manager scan for web part zones and pre-populated web parts
         *  - Detect for field controls - only the metadata behind these can be transformed without an SPFX web part
         *      - Metadata mapping to web part - only some types will be supported
         *  - Using HTML parser deep analysis of the file to map out detected web parts. These are fixed point in the publishing layout.
         *      - This same method could be used to parse HTML fields for inline web parts
         *  - Generate a layout mapping based on analysis
         *  - Validate the Xml prior to output
         *  - Split into molecules of operation for unit testing
         *  - Detect grid system, table or fabric for layout options, needs to be extensible - consider...
         *  
         */

        private ClientContext _siteCollContext;
        private ClientContext _sourceContext;

        private PublishingPageTransformation _mapping;
        private string _defaultFileName = "PageLayoutMapping.xml";

        //TODO: Move to constants class
        const string AvailablePageLayouts = "__PageLayouts";
        const string DefaultPageLayout = "__DefaultPageLayout";
        const string FileRefField = "FileRef";
        const string FileLeafRefField = "FileLeafRef";
        const string PublishingAssociatedContentType = "PublishingAssociatedContentType";
        const string PublishingPageLayoutField = "PublishingPageLayout";
        const string PageLayoutBaseContentTypeId = "0x01010007FF3E057FA8AB4AA42FCB67B453FFC1"; //Page Layout Content Type Id

        private HtmlParser parser;

        /// <summary>
        /// Analyse Page Layouts class constructor
        /// </summary>
        /// <param name="sourceContext">This should be the context of the source web</param>
        /// <param name="logObservers">List of log observers</param>
        public PageLayoutAnalyser(ClientContext sourceContext, IList<ILogObserver> logObservers = null)
        {
            // Register observers
            if (logObservers != null)
            {
                foreach (var observer in logObservers)
                {
                    base.RegisterObserver(observer);
                }
            }

            _mapping = new PublishingPageTransformation();

            _sourceContext = sourceContext;
            EnsureSiteCollectionContext(sourceContext);
            parser = new HtmlParser(new HtmlParserOptions() { IsEmbedded = true }, Configuration.Default.WithDefaultLoader().WithCss());
        }


        /// <summary>
        /// Main entry point into the class to analyse the page layouts
        /// </summary>
        public void AnalyseAll()
        {
            // Determine if ‘default’ layouts for the OOB page layouts
            // When there’s no layout we “generate” a best effort one and store it in cache.Generation can 
            //  be done by looking at the field types and inspecting the layout aspx file. This same generation 
            //  part can be used in point 2 for customers to generate a starting layout mapping file which they then can edit
            // Don't assume that you are in a top level site, you maybe in a sub site

            if (Validate())
            {
                var spPageLayouts = GetAllPageLayouts();

                foreach (ListItem layout in spPageLayouts)
                {
                    AnalysePageLayout(layout);
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pageLayoutMappings"></param>
        /// <param name="pageLayoutItem"></param>
        public void AnalysePageLayout(ListItem pageLayoutItem)
        {

            string assocContentType = pageLayoutItem[PublishingAssociatedContentType].ToString();
            var assocContentTypeParts = assocContentType.Split(new string[] { ";#" }, StringSplitOptions.RemoveEmptyEntries);

            var metadata = GetMetadatafromPageLayoutAssociatedContentType(assocContentTypeParts[1]);
            var extractedHtmlBlocks = ExtractControlsFromPageLayoutHtml(pageLayoutItem);


            var oobPageLayoutDefaults = PublishingDefaults.OOBPageLayouts.FirstOrDefault(o => o.Name == pageLayoutItem.EnsureProperty(i => i.DisplayName));

            var layoutMapping = new PageLayout()
            {
                Name = pageLayoutItem.DisplayName,
                PageHeader = this.CastToEnum<PageLayoutPageHeader>(oobPageLayoutDefaults?.PageHeader),
                PageLayoutTemplate = this.CastToEnum<PageLayoutPageLayoutTemplate>(oobPageLayoutDefaults?.PageLayoutTemplate),
                AssociatedContentType = assocContentTypeParts?[0],
                MetaData = metadata,

                WebParts = extractedHtmlBlocks.WebPartFields.ToArray(),
                WebPartZones = extractedHtmlBlocks.WebPartZones.ToArray(),
                FixedWebParts = extractedHtmlBlocks.FixedWebParts.ToArray()
            };

            SetPageLayoutHeaderFieldDefaults(oobPageLayoutDefaults, layoutMapping);

            // Add to mappings list
            if (_mapping.PageLayouts != null)
            {
                var expandMappings = _mapping.PageLayouts.ToList();
                expandMappings.Add(layoutMapping);
                _mapping.PageLayouts = expandMappings.ToArray();
            }
            else
            {
                _mapping.PageLayouts = new[] { layoutMapping };
            }
        }

        /// <summary>
        /// Sets the page layout header field defaults
        /// </summary>
        /// <param name="oobPageLayoutDefaults"></param>
        /// <param name="layoutMapping"></param>
        private void SetPageLayoutHeaderFieldDefaults(PageLayoutOOBEntity oobPageLayoutDefaults, PageLayout layoutMapping)
        {
            if (layoutMapping.PageHeader == PageLayoutPageHeader.Custom)
            {
                var pageLayoutHeaderFields = PublishingDefaults.PageLayoutHeaderMetadata.Where(o => o.HeaderType == oobPageLayoutDefaults?.PageHeaderType);
                layoutMapping.Header = new Header() { Type = this.CastToEnum<HeaderType>(oobPageLayoutDefaults?.PageHeaderType) };

                List<HeaderField> headerFields = new List<HeaderField>();
                foreach (var field in pageLayoutHeaderFields)
                {
                    headerFields.Add(new HeaderField()
                    {
                        Name = field.FieldName,
                        HeaderProperty = field.FieldHeaderProperty,
                        Functions = field.FieldFunctions
                    });
                }

                layoutMapping.Header.Field = headerFields.ToArray();
            }
        }

        /// <summary>
        /// Perform validation to ensure the source site contains page layouts
        /// </summary>
        public bool Validate()
        {
            if (_sourceContext.Web.IsPublishingWeb())
            {
                return true;
            }

            return false;
        }



        /// <summary>
        /// Determines the page layouts in the current web
        /// </summary>
        public ListItemCollection GetAllPageLayouts()
        {
            var availablePageLayouts = GetPropertyBagValue<string>(_siteCollContext.Web, AvailablePageLayouts, "");
            // If empty then gather all

            var masterPageGallery = _siteCollContext.Web.GetCatalog((int)ListTemplateType.MasterPageCatalog);
            _siteCollContext.Load(masterPageGallery, x => x.RootFolder.ServerRelativeUrl);
            _siteCollContext.ExecuteQueryRetry();

            var query = new CamlQuery();
            // Use query Scope='RecursiveAll' to iterate through sub folders of Master page library because we might have file in folder hierarchy
            // Ensure that we are getting layouts with at least one published version, not hidden layouts
            query.ViewXml =
                $"<View Scope='RecursiveAll'>" +
                    $"<Query>" +
                        $"<Where>" +
                            $"<And>" +
                                $"<And>" +
                                    $"<Geq>" +
                                        $"<FieldRef Name='_UIVersionString'/><Value Type='Text'>1.0</Value>" +
                                    $"</Geq>" +
                                    $"<BeginsWith>" +
                                        $"<FieldRef Name='ContentTypeId'/><Value Type='ContentTypeId'>{PageLayoutBaseContentTypeId}</Value>" +
                                    $"</BeginsWith>" +
                                $"</And>" +
                                $"<Or>" +
                                    $"<Eq>" +
                                        $"<FieldRef Name='PublishingHidden'/><Value Type='Boolean'>0</Value>" +
                                    $"</Eq>" +
                                    $"<IsNull>" +
                                        $"<FieldRef Name='PublishingHidden'/>" +
                                    $"</IsNull>" +
                                $"</Or>" +
                            $"</And>" +
                         $"</Where>" +
                    $"</Query>" +
                    $"<ViewFields>" +
                        $"<FieldRef Name='" + PublishingAssociatedContentType + $"' />" +
                        $"<FieldRef Name='PublishingHidden' />" +
                        $"<FieldRef Name='Title' />" +
                    $"</ViewFields>" +
                  $"</View>";

            var galleryItems = masterPageGallery.GetItems(query);
            _siteCollContext.Load(masterPageGallery);
            _siteCollContext.Load(galleryItems);
            _siteCollContext.Load(galleryItems, i => i.Include(o => o.DisplayName),
                i => i.Include(o => o.File),
                i => i.Include(o => o.File.ServerRelativeUrl));

            _siteCollContext.ExecuteQueryRetry();

            return galleryItems.Count > 0 ? galleryItems : null;

        }

        /// <summary>
        /// Gets the page layout for analysis
        /// </summary>
        public WebPartField[] GetPageLayoutFileWebParts(ListItem pageLayout)
        {

            List<WebPartField> wpFields = new List<WebPartField>();

            File file = pageLayout.File;
            var webPartManager = file.GetLimitedWebPartManager(Microsoft.SharePoint.Client.WebParts.PersonalizationScope.Shared);

            _siteCollContext.Load(webPartManager);
            _siteCollContext.Load(webPartManager.WebParts);
            _siteCollContext.Load(webPartManager.WebParts,
                i => i.Include(o => o.WebPart.Title),
                i => i.Include(o => o.ZoneId),
                i => i.Include(o => o.WebPart));
            _siteCollContext.Load(file);
            _siteCollContext.ExecuteQueryRetry();

            var wps = webPartManager.WebParts;

            foreach (var part in wps)
            {

                var props = part.WebPart.Properties.FieldValues;
                List<WebPartProperty> partProperties = new List<WebPartProperty>();

                foreach (var prop in props)
                {
                    partProperties.Add(new WebPartProperty() { Name = prop.Key, Type = WebPartProperyType.@string });
                }

                wpFields.Add(new WebPartField()
                {
                    Name = part.WebPart.Title,
                    Property = partProperties.ToArray()

                });

            }

            return wpFields.ToArray();
        }


        /// <summary>
        /// Determine the page layout from a publishing page
        /// </summary>
        public void GetPageLayoutFromPublishingPage(ListItem page)
        {
            //Note: ListItemExtensions class contains this logic - reuse.
            //throw new NotImplementedException();


        }

        /// <summary>
        /// Get Metadata mapping from the page layout associated content type
        /// </summary>
        /// <param name="contentTypeId">Id of the content type</param>
        public MetaDataField[] GetMetadatafromPageLayoutAssociatedContentType(string contentTypeId)
        {
            List<MetaDataField> fields = new List<MetaDataField>();

            try
            {

                if (_siteCollContext.Web.ContentTypeExistsById(contentTypeId, true))
                {
                    var cType = _siteCollContext.Web.ContentTypes.GetById(contentTypeId);

                    var spFields = cType.EnsureProperty(o => o.Fields);

                    foreach (var fld in spFields.Where(o => o.Hidden == false))
                    {
                        var ignoreField = PublishingDefaults.IgnoreMetadataFields.Any(o => o == fld.InternalName);
                        var defaultMapping = PublishingDefaults.MetaDataFieldToTargetMappings.FirstOrDefault(o => o.FieldName == fld.InternalName);

                        fields.Add(new MetaDataField()
                        {
                            Name = fld.InternalName,
                            Functions = defaultMapping?.Functions ?? "",
                            TargetFieldName = defaultMapping?.TargetFieldName ?? "",
                            Ignore = ignoreField,
                            IgnoreSpecified = ignoreField
                        });
                    }
                }

            }
            catch (Exception ex)
            {
                LogError(LogStrings.Error_CannotMapMetadataFields, LogStrings.Heading_PageLayoutAnalyser, ex);
            }

            return fields.ToArray();
        }


        /// <summary>
        /// Get fixed web parts defined in the page layout
        /// </summary>
        public FixedWebPart[] GetFixedWebPartsFromZones(ListItem pageLayout)
        {
            /*Plan
             * Scan through the file to find the web parts by the tags
             * Extract and convert to definition
             * Check the TagPrefix and find all the web parts e.g. Register Tagprefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages"
             * Get a list of all the API recognised web parts and perform a delta
             * List of types can be found in the WebParts class in root of project
            */

            List<FixedWebPart> fixedWebParts = new List<FixedWebPart>();
            const string TagPrefix = "WebPartPages";

            pageLayout.EnsureProperties(o => o.File, o => o.File.ServerRelativeUrl);
            var fileUrl = pageLayout.File.ServerRelativeUrl;

            var fileHtml = _siteCollContext.Web.GetFileAsString(fileUrl);

            using (var document = this.parser.Parse(fileHtml))
            {

                var webParts = document.All.Where(o => o.TagName.Contains(TagPrefix)).ToArray();

                for (var i = 0; i < webParts.Count(); i++)
                {
                    fixedWebParts.Add(new FixedWebPart()
                    {
                        Type = "",
                        Column = "",
                        Row = "",
                        Order = ""
                    });
                }

            }

            return fixedWebParts.ToArray();

        }

        /// <summary>
        /// This method analyses the Html strcuture to determine layout
        /// </summary>
        public void ExtractLayoutFromHtmlStructure()
        {
            /*Plan
             * Scan through the file to plot the 
             * - Determine if a grid system, classic, fabric or Html structure is in use
             * - Work out the location of the web part in relation to the grid system
            */
        }

        /// <summary>
        /// Scan through the file to find the TagPrefixes in ASPX Header
        /// </summary>
        /// <param name="pageLayout"></param>
        /// <returns>
        ///     List<Tuple<string, string>>
        ///     Item1 = tagprefix
        ///     Item2 = Namespace
        /// </returns>
        public List<Tuple<string, string>> ExtractWebPartPrefixesFromNamespaces(ListItem pageLayout)
        {
            var tagPrefixes = new List<Tuple<string, string>>();

            pageLayout.EnsureProperties(o => o.File, o => o.File.ServerRelativeUrl);
            var fileUrl = pageLayout.File.ServerRelativeUrl;

            var fileHtml = _siteCollContext.Web.GetFileAsString(fileUrl);

            using (var document = this.parser.Parse(fileHtml))
            {
                Regex regex = new Regex("&lt;%@(.*?)%&gt;", RegexOptions.IgnoreCase | RegexOptions.Multiline);
                var aspxHeader = document.All.Where(o => o.TagName == "HTML").FirstOrDefault();
                var results = regex.Matches(aspxHeader?.InnerHtml);

                StringBuilder blockHtml = new StringBuilder();
                foreach (var match in results)
                {
                    var matchString = match.ToString().Replace("&lt;%@ ", "<").Replace("%&gt;", " />");
                    blockHtml.AppendLine(matchString);
                }

                var fullBlock = blockHtml.ToString();
                using (var subDocument = this.parser.Parse(fullBlock))
                {
                    var registers = subDocument.All.Where(o => o.TagName == "REGISTER");

                    foreach (var register in registers)
                    {
                        var prefix = register.GetAttribute("Tagprefix");
                        var nameSpace = register.GetAttribute("Namespace");
                        tagPrefixes.Add(new Tuple<string, string>(prefix, nameSpace));
                    }

                }

            }

            return tagPrefixes;
        }

        /// <summary>
        /// Extract the web parts from the page layout HTML outside of web part zones
        /// </summary>
        public ExtractedHtmlBlocksEntity ExtractControlsFromPageLayoutHtml(ListItem pageLayout)
        {
            /*Plan
             * Scan through the file to find the web parts by the tags
             * Extract and convert to definition 
            */

            ExtractedHtmlBlocksEntity extractedHtmlBlocks = new ExtractedHtmlBlocksEntity();

            // Data from SharePoint
            pageLayout.EnsureProperties(o => o.File, o => o.File.ServerRelativeUrl);
            var fileUrl = pageLayout.File.ServerRelativeUrl;
            var fileHtml = _siteCollContext.Web.GetFileAsString(fileUrl);

            using (var document = this.parser.Parse(fileHtml))
            {

                // Item 1 - WebPart Name, Item 2 - Full assembly reference
                List<Tuple<string, string>> possibleWebPartsUsed = new List<Tuple<string, string>>();
                List<IEnumerable<IElement>> multipleTagFinds = new List<IEnumerable<IElement>>();

                //List of all the assembly references and prefixes in the page
                List<Tuple<string, string>> prefixesAndNameSpaces = ExtractWebPartPrefixesFromNamespaces(pageLayout);

                // Determine the possible web parts from the page from the namespaces used in the aspx header
                prefixesAndNameSpaces.ForEach(p =>
                {
                    var possibleParts = WebParts.GetListOfWebParts(p.Item2);
                    foreach (var part in possibleParts)
                    {
                        var webPartName = part.Substring(0, part.IndexOf(",")).Replace($"{p.Item2}.", "");
                        possibleWebPartsUsed.Add(new Tuple<string, string>(webPartName, part));
                    }
                });

                // Cycle through all the nodes in the document
                foreach (var docNode in document.All)
                {
                    foreach (var prefixAndNameSpace in prefixesAndNameSpaces)
                    {
                        if (docNode.TagName.Contains(prefixAndNameSpace.Item1.ToUpper()))
                        {

                            // Expand, as this may contain many elements
                            //foreach (var control in tagFind)
                            //{

                            var attributes = docNode.Attributes;
                            
                            if (attributes.Any(o => o.Name == "fieldname"))
                            {

                                var fieldName = attributes["fieldname"].Value;

                                //DeDup - Some controls can be inside an edit panel
                                if (!extractedHtmlBlocks.WebPartFields.Any(o => o.Name == fieldName))
                                {
                                    extractedHtmlBlocks.WebPartFields.Add(new WebPartField()
                                    {
                                        Name = fieldName,
                                        TargetWebPart = "",
                                        Row = "",
                                        Column = ""

                                    });
                                }
                            }

                            if (docNode.TagName.Contains("WEBPARTZONE"))
                            {

                                extractedHtmlBlocks.WebPartZones.Add(new WebPartZone()
                                {
                                    ZoneId = docNode.Id,
                                    Column = "",
                                    Row = ""
                                    //ZoneIndex = control. // TODO: Is this used?
                                });
                            }

                            //Fixed web part zone
                            //This should only find one match
                            var matchedParts = possibleWebPartsUsed.Where(o => o.Item1.ToUpper() == docNode.TagName.Replace($"{prefixAndNameSpace.Item1.ToUpper()}:", ""));

                            if (matchedParts.Any())
                            {
                                var match = matchedParts.FirstOrDefault();
                                if(match != default(Tuple<string, string>))
                                {
                                    //Process Child properties
                                    List<FixedWebPartProperty> fixedProperties = new List<FixedWebPartProperty>();
                                    if (docNode.HasChildNodes && docNode.FirstElementChild.HasChildNodes) {
                                        var childProperties = docNode.FirstElementChild.ChildNodes;
                                        foreach(var childProp in childProperties) {

                                            if (childProp.NodeName != "#text")
                                            {
                                                var stronglyTypedChild = (IElement)childProp;
                                                var content = !string.IsNullOrEmpty(childProp.TextContent) ? childProp.TextContent : stronglyTypedChild.InnerHtml;

                                                fixedProperties.Add(new FixedWebPartProperty()
                                                {
                                                    Name = stronglyTypedChild.NodeName,
                                                    Type = WebPartProperyType.@string,
                                                    Value = System.Web.HttpUtility.HtmlEncode(content)
                                                });
                                            }
                                        }
                                    }

                                    extractedHtmlBlocks.FixedWebParts.Add(new FixedWebPart()
                                    {
                                        Column = "",
                                        Row = "",
                                        Type = match.Item2,
                                        Property = fixedProperties.ToArray()
                                    });
                                }
                            }
                        }
                    }
                }


                //foreach (var prefixAndNameSpace in prefixesAndNameSpaces)
                //{
                //    multipleTagFinds.Add(document.All.Where(o => o.TagName.Contains(prefixAndNameSpace.Item1.ToUpper())));

                //    // Determine the possible web parts from the page from the namespaces used in the aspx header
                //    var possibleParts = WebParts.GetListOfWebParts(prefixAndNameSpace.Item2);
                //    foreach(var part in possibleParts)
                //    {
                //        var webPartName = part.Substring(0, part.IndexOf(",")).Replace(prefixAndNameSpace.Item2, "");
                //        possibleWebPartsUsed.Add(new Tuple<string, string>(webPartName, part));
                //    }
                //}

                //// Bit of a bad name, just getting it working, this refers to all sharepoint controls including web parts, zones and field controls.
                //foreach (var tagFind in multipleTagFinds)
                //{


                //}

            }

            return extractedHtmlBlocks;

        }

        /// <summary>
        /// Extract the web parts from the page layout HTML outside of web part zones
        /// </summary>
        public WebPartZone[] ExtractWebPartZonesFromPageLayoutHtml(ListItem pageLayout)
        {
            /*Plan
             * Scan through the file to find the web parts by the tags
             * Extract and convert to definition 
            */
            List<WebPartZone> zones = new List<WebPartZone>();

            pageLayout.EnsureProperties(o => o.File, o => o.File.ServerRelativeUrl);
            var fileUrl = pageLayout.File.ServerRelativeUrl;

            var fileHtml = _siteCollContext.Web.GetFileAsString(fileUrl);

            using (var document = this.parser.Parse(fileHtml))
            {
                //TODO: Add further processing to find if the tags are in a grid system

                var webPartZones = document.All.Where(o => o.TagName.Contains("WEBPARTZONE")).ToArray();

                for (var i = 0; i < webPartZones.Count(); i++)
                {
                    zones.Add(new WebPartZone()
                    {
                        ZoneId = webPartZones[i].Id,
                        Column = "",
                        Row = "",
                        ZoneIndex = $"{i}" // TODO: Is this used?
                    });
                }

            }

            return zones.ToArray();

        }

        /// <summary>
        /// Generate the mapping file to output from the analysis
        /// </summary>
        public string GenerateMappingFile()
        {
            try
            {
                XmlSerializer xmlMapping = new XmlSerializer(typeof(PublishingPageTransformation));

                var mappingFileName = _defaultFileName;

                using (StreamWriter sw = new StreamWriter(mappingFileName, false))
                {
                    xmlMapping.Serialize(sw, _mapping);
                }

                var xmlMappingFileLocation = $"{ Environment.CurrentDirectory }\\{ mappingFileName}";
                LogInfo($"{LogStrings.XmlMappingSavedAs}: {xmlMappingFileLocation}");

                return xmlMappingFileLocation;

            }
            catch (Exception ex)
            {
                var message = string.Format(LogStrings.Error_CannotWriteToXmlFile, ex.Message, ex.StackTrace);
                Console.WriteLine(message);
                LogError(message, LogStrings.Heading_PageLayoutAnalyser, ex);
            }

            return string.Empty;
        }


        #region Helpers

        /// <summary>
        /// Ensures that we have context of the source site collection
        /// </summary>
        /// <param name="context"></param>
        public void EnsureSiteCollectionContext(ClientContext context)
        {
            try
            {
                if (context.Web.IsSubSite())
                {
                    string siteCollectionUrl = context.Site.EnsureProperty(o => o.Url);
                    _siteCollContext = context.Clone(siteCollectionUrl);
                }
                else
                {
                    _siteCollContext = context;
                }
            }
            catch (Exception ex)
            {
                LogError(LogStrings.Error_CannotGetSiteCollContext, LogStrings.Heading_PageLayoutAnalyser, ex);
            }
        }

        /// <summary>
        /// Gets property bag value
        /// </summary>
        /// <typeparam name="T">Cast to type of</typeparam>
        /// <param name="web">Current Web</param>
        /// <param name="key">KeyValue Pair - Key</param>
        /// <param name="defaultValue">Default Value</param>
        /// <returns></returns>
        private static T GetPropertyBagValue<T>(Web web, string key, T defaultValue)
        {
            //TODO: Add to helpers class - source from Publishing Analyser

            web.EnsureProperties(p => p.AllProperties);

            if (web.AllProperties.FieldValues.ContainsKey(key))
            {
                return (T)web.AllProperties.FieldValues[key];
            }
            else
            {
                return defaultValue;
            }
        }


        /// <summary>
        /// Cast a string to enum value
        /// </summary>
        /// <typeparam name="T">Enum Type</typeparam>
        /// <param name="enumString">string value</param>
        /// <returns></returns>
        private T CastToEnum<T>(string enumString)
        {
            if (!string.IsNullOrEmpty(enumString))
            {
                try
                {

                    return (T)Enum.Parse(typeof(T), enumString, true);

                }
                catch (Exception ex)
                {
                    LogError(LogStrings.Error_CannotCastToEnum, LogStrings.Heading_PageLayoutAnalyser, ex);
                }
            }

            return default(T);
        }

        #endregion
    }

    /// <summary>
    /// Simple entity for the extracted blocks of data
    /// </summary>
    public class ExtractedHtmlBlocksEntity
    {
        public ExtractedHtmlBlocksEntity()
        {
            WebPartFields = new List<WebPartField>();
            WebPartZones = new List<WebPartZone>();
            FixedWebParts = new List<FixedWebPart>();
        }

        public List<WebPartField> WebPartFields { get; set; }
        public List<WebPartZone> WebPartZones { get; set; }
        public List<FixedWebPart> FixedWebParts { get; set; }
    }
}
