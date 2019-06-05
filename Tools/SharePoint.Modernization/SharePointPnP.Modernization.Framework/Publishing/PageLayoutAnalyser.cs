using Microsoft.SharePoint.Client;
using SharePointPnP.Modernization.Framework.Telemetry;
using SharePointPnP.Modernization.Framework.Transform;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using AngleSharp;
using AngleSharp.Dom;
using AngleSharp.Parser.Html;
using File = Microsoft.SharePoint.Client.File;
using SharePointPnP.Modernization.Framework.Publishing.Layouts;
using System.Text.RegularExpressions;
using SharePointPnP.Modernization.Framework.Extensions;

namespace SharePointPnP.Modernization.Framework.Publishing
{
    public class PageLayoutAnalyser : BaseTransform
    {
        /// <summary>
        /// Simple entity for the extracted blocks of data
        /// </summary>
        internal class ExtractedHtmlBlocksEntity
        {
            internal ExtractedHtmlBlocksEntity()
            {
                WebPartFields = new List<WebPartField>();
                WebPartZones = new List<WebPartZone>();
                FixedWebParts = new List<FixedWebPart>();
            }

            internal List<WebPartField> WebPartFields { get; set; }
            internal List<WebPartZone> WebPartZones { get; set; }
            internal List<FixedWebPart> FixedWebParts { get; set; }
        }

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

        private PublishingPageTransformation _mapping;
        private string _defaultFileName = "PageLayoutMapping.xml";
        private HtmlParser parser;
        private Dictionary<string, FieldCollection> _contentTypeFieldCache;
        private Dictionary<string, string> _pageLayoutFileCache;

        #region Construction
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
            _contentTypeFieldCache = new Dictionary<string, FieldCollection>();
            _pageLayoutFileCache = new Dictionary<string, string>();

            EnsureSiteCollectionContext(sourceContext);

            //parser = new HtmlParser(new HtmlParserOptions() { IsEmbedded = true,}, Configuration.Default.WithDefaultLoader().WithCss());
            parser = new HtmlParser();
        }
        #endregion

        #region Public interface
        /// <summary>
        /// Main entry point into the class to analyse the page layouts
        /// </summary>
        /// <param name="skipOOBPageLayouts">Skip OOB page layouts</param>
        public void AnalyseAll(bool skipOOBPageLayouts = false)
        {
            // Determine if ‘default’ layouts for the OOB page layouts
            // When there’s no layout we “generate” a best effort one and store it in cache. Generation can 
            // be done by looking at the field types and inspecting the layout aspx file. This same generation 
            // part can be used in point 2 for customers to generate a starting layout mapping file which they then can edit
            // Don't assume that you are in a top level site, you maybe in a sub site

            var spPageLayouts = GetAllPageLayouts();

            if (spPageLayouts != null)
            {
                foreach (ListItem layout in spPageLayouts)
                {
                    try
                    {
                        if (skipOOBPageLayouts)
                        {
                            // Check if this is an OOB page layout and skip if so
                            string pageLayoutName = Path.GetFileNameWithoutExtension(layout[Constants.FileLeafRefField].ToString());                            
                            var oobPageLayoutFile = Constants.OobPublishingPageLayouts.Any(o => o.Equals(pageLayoutName, StringComparison.InvariantCultureIgnoreCase));
                            if (oobPageLayoutFile)
                            {
                                LogInfo(String.Format(LogStrings.OOBPageLayoutSkipped, pageLayoutName), LogStrings.Heading_PageLayoutAnalyser);
                                continue;
                            }
                        }

                        AnalysePageLayout(layout);
                    }
                    catch (Exception ex)
                    {
                        LogError(LogStrings.Error_CannotProcessPageLayoutAnalyseAll, LogStrings.Heading_PageLayoutAnalyser, ex);
                    }

                }
            }
        }

        /// <summary>
        /// Analyses a single page layout from a provided file
        /// </summary>
        /// <param name="pageLayoutItem">Page layout list item</param>
        public PageLayout AnalysePageLayout(ListItem pageLayoutItem)
        {
            try
            {
                // Get the associated page layout content type
                string assocContentType = pageLayoutItem[Constants.PublishingAssociatedContentTypeField].ToString();
                var assocContentTypeParts = assocContentType.Split(new string[] { ";#" }, StringSplitOptions.RemoveEmptyEntries);

                // Load content type fields in memory once
                var contentTypeFields = LoadContentTypeFields(assocContentTypeParts[1]);

                // Extact page header
                var extractedHeader = ExtractPageHeaderFromPageLayoutAssociatedContentType(contentTypeFields);

                // Analyze the pagelayout file content
                var extractedHtmlBlocks = ExtractControlsFromPageLayoutHtml(pageLayoutItem);
                extractedHtmlBlocks.WebPartFields = CleanExtractedWebPartFields(extractedHtmlBlocks.WebPartFields, contentTypeFields);

                // Detect the fields that will become metadata in the target site
                var extractedMetaDataFields = ExtractMetaDataFromPageLayoutAssociatedContentType(contentTypeFields, extractedHtmlBlocks.WebPartFields, extractedHeader);

                var metaData = new MetaData
                {                    
                    Field = extractedMetaDataFields
                };

                // Combine all data to a single PageLayout mapping
                var layoutMapping = new PageLayout()
                {
                    // Display name of the page layout
                    Name = Path.GetFileNameWithoutExtension(pageLayoutItem[Constants.FileLeafRefField].ToString()),
                    // Default to no page header for now
                    PageHeader = extractedHeader != null ? PageLayoutPageHeader.Custom : PageLayoutPageHeader.None,
                    // Default to autodetect layout model
                    PageLayoutTemplate = PageLayoutPageLayoutTemplate.AutoDetect,
                    // The content type to be used on the target modern page
                    AssociatedContentType = "",
                    // Set the header details (if any)
                    Header = extractedHeader,
                    // Fields that will become metadata fields for the target page
                    MetaData = metaData,
                    // Fields that will become web parts on the target page
                    WebParts = extractedHtmlBlocks.WebPartFields.Count > 0 ? extractedHtmlBlocks.WebPartFields.ToArray() : null,
                    // Web part zones that can hold zero or more web parts
                    WebPartZones = extractedHtmlBlocks.WebPartZones.Count > 0 ? extractedHtmlBlocks.WebPartZones.ToArray() : null,
                    // Fixed web parts, this are web parts which are 'hardcoded' in the pagelayout aspx file
                    FixedWebParts = extractedHtmlBlocks.FixedWebParts.Count > 0 ? extractedHtmlBlocks.FixedWebParts?.ToArray() : null,
                };

                // Add to mappings list
                if (_mapping.PageLayouts != null)
                {
                    var expandMappings = _mapping.PageLayouts.ToList();

                    // Prevent duplicate references to the same page layout
                    if (!expandMappings.Any(o => o.Name == layoutMapping.Name))
                    {
                        expandMappings.Add(layoutMapping);
                    }

                    _mapping.PageLayouts = expandMappings.ToArray();
                }
                else
                {
                    _mapping.PageLayouts = new[] { layoutMapping };
                }

                return layoutMapping;

            }catch(Exception ex)
            {
                LogError(LogStrings.Error_CannotProcessPageLayoutAnalyse, LogStrings.Heading_PageLayoutAnalyser, ex);
            }

            return null;
        }

        /// <summary>
        /// Determine the page layout from a publishing page
        /// </summary>
        /// <param name="publishingPage">Publishing page to analyze the page layout for</param>
        public PageLayout AnalysePageLayoutFromPublishingPage(ListItem publishingPage)
        {
            //Note: ListItemExtensions class contains this logic - reuse.
            //TODO: Make more defensive, this could represent the wrong item 
            var pageLayoutFileUrl = publishingPage.PageLayoutFile();

            if (!string.IsNullOrEmpty(pageLayoutFileUrl))
            {
                Uri uri = new Uri(pageLayoutFileUrl);
                var host = $"{uri.Scheme}://{uri.Host}";
                var path = pageLayoutFileUrl.Replace(host, "");

                var file = _siteCollContext.Web.GetFileByServerRelativeUrl(path);
                _siteCollContext.Load(file, o => o.ListItemAllFields);
                _siteCollContext.ExecuteQueryRetry();

                return AnalysePageLayout(file.ListItemAllFields);
            }
            else
            {
                // Add logging here
                throw new ArgumentNullException("Page layout could not be determined by the publishing page");
            }
        }

        /// <summary>
        /// Generate the mapping file to output from the analysis
        /// </summary>
        /// <returns>Mapping file fully qualified path</returns>
        public string GenerateMappingFile()
        {
            return GenerateMappingFile(Environment.CurrentDirectory, _defaultFileName);
        }

        /// <summary>
        /// Generate the mapping file to output from the analysis
        /// </summary>
        /// <param name="folder">Folder to generate the file in</param>
        /// <returns>Mapping file fully qualified path</returns>
        public string GenerateMappingFile(string folder)
        {
            return GenerateMappingFile(folder, _defaultFileName);
        }

        /// <summary>
        /// Generate the mapping file to output from the analysis
        /// </summary>
        /// <param name="folder">Folder to generate the file in</param>
        /// <param name="fileName">name of the mapping file</param>
        /// <returns>Mapping file fully qualified path</returns>
        public string GenerateMappingFile(string folder, string fileName)
        {
            try
            {
                XmlSerializer xmlMapping = new XmlSerializer(typeof(PublishingPageTransformation));

                var mappingFileName = $"{folder}\\{fileName}";
                using (StreamWriter sw = new StreamWriter(mappingFileName, false))
                {
                    xmlMapping.Serialize(sw, _mapping);
                }

                LogInfo($"{LogStrings.XmlMappingSavedAs}: {mappingFileName}");
                return mappingFileName;
            }
            catch (Exception ex)
            {
                var message = string.Format(LogStrings.Error_CannotWriteToXmlFile, ex.Message, ex.StackTrace);
                Console.WriteLine(message);
                LogError(message, LogStrings.Heading_PageLayoutAnalyser, ex);
            }

            return string.Empty;
        }
        #endregion

        #region Internal methods
        /// <summary>
        /// Determines the page layouts in the current web
        /// </summary>
        internal ListItemCollection GetAllPageLayouts()
        {
            var masterPageGallery = _siteCollContext.Web.GetCatalog((int)ListTemplateType.MasterPageCatalog);
            _siteCollContext.Load(masterPageGallery, x => x.RootFolder.ServerRelativeUrl);

            var query = new CamlQuery
            {
                // Use query Scope='RecursiveAll' to iterate through sub folders of Master page library because we might have file in folder hierarchy
                // Ensure that we are getting layouts with at least one published version, not hidden layouts
                ViewXml =
                $"<View Scope='RecursiveAll'>" +
                    $"<Query>" +
                        $"<Where>" +
                            $"<And>" +
                                $"<Contains>" +
                                    $"<FieldRef Name='File_x0020_Type'/><Value Type='Text'>aspx</Value>" +
                                $"</Contains>" +
                                $"<And>" +
                                    $"<And>" +
                                        $"<Geq>" +
                                            $"<FieldRef Name='_UIVersionString'/><Value Type='Text'>1.0</Value>" +
                                        $"</Geq>" +
                                        $"<BeginsWith>" +
                                            $"<FieldRef Name='ContentTypeId'/><Value Type='ContentTypeId'>{Constants.PageLayoutBaseContentTypeId}</Value>" +
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
                            $"</And>" +
                         $"</Where>" +
                    $"</Query>" +
                    $"<ViewFields>" +
                        $"<FieldRef Name='{Constants.PublishingAssociatedContentTypeField}' />" +
                        $"<FieldRef Name='PublishingHidden' />" +
                        $"<FieldRef Name='Title' />" +
                    $"</ViewFields>" +
                  $"</View>"
            };

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
        /// Get Metadata mapping from the page layout associated content type
        /// </summary>
        /// <param name="contentTypeId">Id of the content type</param>
        internal MetaDataField[] ExtractMetaDataFromPageLayoutAssociatedContentType(FieldCollection spFields, List<WebPartField> webPartFields, Header extractedHeader)
        {
            List<MetaDataField> fields = new List<MetaDataField>();

            // Get unique field types for which we've defined web part mapping defaults
            List<string> fieldTypesToSkip = new List<string>();
            foreach(var defaultWebPartField in PublishingDefaults.WebPartFieldProperties)
            {
                if (!fieldTypesToSkip.Contains(defaultWebPartField.FieldType))
                {
                    fieldTypesToSkip.Add(defaultWebPartField.FieldType);
                }
            }

            // Skip hidden fields by default
            foreach (var spField in spFields.Where(o => o.Hidden == false))
            {
                if (!PublishingDefaults.IgnoreMetadataFields.Any(o => o.Equals(spField.InternalName, StringComparison.InvariantCultureIgnoreCase)))
                {
                    // Was this field already defined as a field that will be mapped to a web part? If so it can't be a metadata field
                    if (webPartFields.Where(p => p.Name.Equals(spField.InternalName, StringComparison.InvariantCultureIgnoreCase)).Any())
                    {
                        continue;
                    }

                    // Was this field already defined as a field that will be mapped to a header property? If so it can't be a metadata field
                    if (extractedHeader != null)
                    {
                        if (extractedHeader.Field.Where(p => p.Name.Equals(spField.InternalName, StringComparison.InvariantCultureIgnoreCase)).Any())
                        {
                            continue;
                        }
                    }

                    // Any field of a type that by default has as target a web part typically is not meant as metadata field for users
                    if (fieldTypesToSkip.Contains(spField.TypeAsString))
                    {
                        continue;
                    }

                    // Load the default mapping information for this field
                    var defaultMapping = PublishingDefaults.MetaDataFieldToTargetMappings.FirstOrDefault(o => o.FieldName.Equals(spField.InternalName, StringComparison.InvariantCultureIgnoreCase));
                    fields.Add(new MetaDataField()
                    {
                        Name = spField.InternalName,
                        Functions = defaultMapping?.Functions ?? "",
                        TargetFieldName = defaultMapping?.TargetFieldName ?? "",
                    });
                }
            }

            return fields.ToArray();
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
        internal List<Tuple<string, string>> ExtractWebPartPrefixesFromNamespaces(ListItem pageLayout)
        {
            var tagPrefixes = new List<Tuple<string, string>>();

            pageLayout.EnsureProperties(o => o.File, o => o.File.ServerRelativeUrl);
            var fileUrl = pageLayout.File.ServerRelativeUrl;
            string fileHtml = LoadPageLayoutFile(fileUrl);

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
        internal ExtractedHtmlBlocksEntity ExtractControlsFromPageLayoutHtml(ListItem pageLayout)
        {
            /*Plan
             * Scan through the file to find the web parts by the tags
             * Extract and convert to definition 
            */

            ExtractedHtmlBlocksEntity extractedHtmlBlocks = new ExtractedHtmlBlocksEntity();

            // Data from SharePoint
            pageLayout.EnsureProperties(o => o.File, o => o.File.ServerRelativeUrl);
            var fileUrl = pageLayout.File.ServerRelativeUrl;
            var fileHtml = LoadPageLayoutFile(fileUrl);

            // replace cdata tags to 'fool' AngleSharp
            fileHtml = fileHtml.Replace("<![CDATA[", "<encodeddata>");
            fileHtml = fileHtml.Replace("]]>", "</encodeddata>");

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
                                    List<WebPartProperty> webPartProperties = new List<WebPartProperty>();

                                    foreach (var attr in attributes)
                                    {
                                        // This might need a filter

                                        webPartProperties.Add(new WebPartProperty()
                                        {
                                            Name = attr.Name,
                                            Type = WebPartProperyType.@string,
                                            Functions = "" // Need defaults here
                                        });
                                    }

                                    extractedHtmlBlocks.WebPartFields.Add(new WebPartField()
                                    {
                                        Name = fieldName,
                                        TargetWebPart = "",
                                        Row = 1,
                                        //RowSpecified = true,
                                        Column = 1,
                                        //ColumnSpecified = true,
                                        Property = webPartProperties.ToArray()
                                    });
                                }
                            }

                            if (docNode.TagName.Contains("WEBPARTZONE"))
                            {

                                extractedHtmlBlocks.WebPartZones.Add(new WebPartZone()
                                {
                                    ZoneId = docNode.Id,
                                    Column = 1,
                                    Row = 1,
                                    //ZoneIndex = control. // TODO: Is this used?
                                });
                            }

                            //Fixed web part zone
                            //This should only find one match
                            var matchedParts = possibleWebPartsUsed.Where(o => o.Item1.ToUpper() == docNode.TagName.Replace($"{prefixAndNameSpace.Item1.ToUpper()}:", ""));

                            if (matchedParts.Any())
                            {
                                var match = matchedParts.FirstOrDefault();
                                if (match != default(Tuple<string, string>))
                                {
                                    //Process Child properties
                                    List<FixedWebPartProperty> fixedProperties = new List<FixedWebPartProperty>();
                                    if (docNode.HasChildNodes && docNode.FirstElementChild != null && docNode.FirstElementChild.HasChildNodes)
                                    {
                                        var childProperties = docNode.FirstElementChild.ChildNodes;
                                        foreach (var childProp in childProperties)
                                        {

                                            if (childProp.NodeName != "#text")
                                            {
                                                var stronglyTypedChild = (IElement)childProp;
                                                //var content = !string.IsNullOrEmpty(childProp.TextContent) ? childProp.TextContent : stronglyTypedChild.InnerHtml;
                                                var content = stronglyTypedChild.InnerHtml;

                                                fixedProperties.Add(new FixedWebPartProperty()
                                                {
                                                    Name = stronglyTypedChild.NodeName,
                                                    Type = WebPartProperyType.@string,
                                                    Value = EncodingAndCleanUpContent(content)
                                                });
                                            }
                                        }
                                    }
                                    else
                                    {
                                        // Another scenario where there are no child nodes, just attributes
                                        foreach (var attr in attributes)
                                        {
                                            // This might need a filter

                                            fixedProperties.Add(new FixedWebPartProperty()
                                            {
                                                Name = attr.Name,
                                                Type = WebPartProperyType.@string,
                                                Value = attr.Value
                                            });
                                        }
                                    }

                                    extractedHtmlBlocks.FixedWebParts.Add(new FixedWebPart()
                                    {
                                        Column = 1,
                                        //ColumnSpecified = true,
                                        Row = 1,
                                        //RowSpecified = true,
                                        Type = match.Item2,
                                        Property = fixedProperties.ToArray()
                                    });
                                }
                            }
                        }
                    }
                }

            }

            return extractedHtmlBlocks;
        }

        /// <summary>
        /// Cleans and encodes content data
        /// </summary>
        /// <param name="content">web part value</param>
        /// <returns></returns>
        internal string EncodingAndCleanUpContent(string content)
        {
            if (string.IsNullOrEmpty(content))
            {
                return "";
            }

            if (content.Contains("encodeddata>"))
            {
                // Drop the 'fake' tags again...we'll the XML serializer deal with the encoding work
                content = content.Replace("<encodeddata>", "").Replace("</encodeddata>", "");
                return content;
            }
            else
            {
                return System.Web.HttpUtility.HtmlEncode(content);
            }
        }

        #endregion

        #region Helper methods
        private string LoadPageLayoutFile(string fileUrl)
        {
            // Try to get from cache
            if (_pageLayoutFileCache.TryGetValue(fileUrl, out string fileContentsFromCache))
            {
                return fileContentsFromCache;
            }

            // Load from SharePoint
            string fileContents = _siteCollContext.Web.GetFileByServerRelativeUrlAsString(fileUrl);

            // Store in cache
            _pageLayoutFileCache.Add(fileUrl, fileContents);

            return fileContents;
        }

        /// <summary>
        /// Loads the content type fields
        /// </summary>
        /// <param name="contentTypeId"></param>
        /// <returns></returns>
        private FieldCollection LoadContentTypeFields(string contentTypeId)
        {
            try
            {
                // Try loading from cache first
                if (_contentTypeFieldCache.TryGetValue(contentTypeId, out FieldCollection spFieldsFromCache))
                {
                    return spFieldsFromCache;
                }

                var cType = _siteCollContext.Web.ContentTypes.GetById(contentTypeId);
                var spFields = cType.EnsureProperty(o => o.Fields);

                // Add to cache
                _contentTypeFieldCache.Add(contentTypeId, spFields);

                return spFields;
            }
            catch (Exception ex)
            {
                LogError(LogStrings.Error_CannotMapMetadataFields, LogStrings.Heading_PageLayoutAnalyser, ex);
                throw;
            }
        }

        /// <summary>
        /// Perform cleanup of web part fields
        /// </summary>
        /// <param name="webPartFields">List of extracted web parts</param>
        /// <param name="spFields">Collection of fields</param>
        /// <returns></returns>
        private List<WebPartField> CleanExtractedWebPartFields(List<WebPartField> webPartFields, FieldCollection spFields)
        {
            List<WebPartField> cleanedWebPartFields = new List<WebPartField>();

            foreach (var webPartField in webPartFields)
            {
                if (PublishingDefaults.IgnoreWebPartFieldControls.Contains(webPartField.Name))
                {
                    // This is field we're ignoring as it's not meant to be translated into a web part on the modern page
                    continue;
                }

                Guid fieldId = Guid.Empty;
                Guid.TryParse(webPartField.Name, out fieldId);

                // Find the field, we'll use the field's type to get the 'default' transformation behaviour
                var spField = spFields.Where(p => p.StaticName.Equals(webPartField.Name, StringComparison.InvariantCultureIgnoreCase) || p.Id == fieldId).FirstOrDefault();
                if (spField != null)
                {
                    var webPartFieldDefaults = PublishingDefaults.WebPartFieldProperties.Where(p => p.FieldType.Equals(spField.TypeAsString));
                    if (webPartFieldDefaults.Any())
                    {
                        // Copy basic fields
                        WebPartField wpf = new WebPartField()
                        {
                            Name = spField.StaticName,
                            Row = webPartField.Row,
                            //RowSpecified = webPartField.RowSpecified,
                            Column = webPartField.Column,
                            //ColumnSpecified = webPartField.ColumnSpecified,
                            TargetWebPart = webPartFieldDefaults.First().TargetWebPart,
                        };

                        if(fieldId != Guid.Empty)
                        {
                            wpf.FieldId = fieldId.ToString();
                        }

                        // Copy the default target web part properties
                        var properties = PublishingDefaults.WebPartFieldProperties.Where(p => p.FieldType.Equals(spField.TypeAsString));
                        if (properties.Any())
                        {
                            List<WebPartProperty> webPartProperties = new List<WebPartProperty>();
                            foreach (var property in properties)
                            {
                                webPartProperties.Add(new WebPartProperty()
                                {
                                    Name = property.Name,
                                    Type = this.CastToEnum<WebPartProperyType>(property.Type),
                                    Functions = property.Functions,
                                });
                            }

                            wpf.Property = webPartProperties.ToArray();
                        }

                        cleanedWebPartFields.Add(wpf);
                    }
                }
            }

            return cleanedWebPartFields;
        }

        /// <summary>
        /// Sets the page layout header field defaults
        /// </summary>
        /// <param name="oobPageLayoutDefaults"></param>
        /// <param name="layoutMapping"></param>
        private Header ExtractPageHeaderFromPageLayoutAssociatedContentType(FieldCollection spFields)
        {
            // If we've a publishing rollup image then let's try to use that as page header image...at conversion time we'll still switch back to no header in case there
            // was no publishing rollup image set at content level
            if (spFields.Where(p=>p.InternalName.Equals("PublishingRollupImage", StringComparison.InvariantCultureIgnoreCase)).Any())
            {
                var pageLayoutHeaderFields = PublishingDefaults.PageLayoutHeaderMetadata.Where(o => o.Type.Equals("FullWidthImage", StringComparison.InvariantCultureIgnoreCase));
                var header = new Header() {
                    Type = HeaderType.FullWidthImage,
                    Alignment = this.CastToEnum<HeaderAlignment>(pageLayoutHeaderFields.First().Alignment),
                    ShowPublishedDate = pageLayoutHeaderFields.First().ShowPublishedDate,
                    ShowPublishedDateSpecified = true,
                };

                List<HeaderField> headerFields = new List<HeaderField>();
                foreach (var field in pageLayoutHeaderFields)
                {
                    headerFields.Add(new HeaderField()
                    {
                        Name = field.Name,
                        HeaderProperty = this.CastToEnum<HeaderFieldHeaderProperty>(field.HeaderProperty),
                        Functions = field.Functions
                    });
                }

                header.Field = headerFields.ToArray();

                return header;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Ensures that we have context of the source site collection
        /// </summary>
        /// <param name="context">Connection to SharePoint</param>
        private void EnsureSiteCollectionContext(ClientContext context)
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
}