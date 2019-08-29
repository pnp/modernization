using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Pages;
using OfficeDevPnP.Core.Utilities;
using SharePointPnP.Modernization.Framework.Cache;
using SharePointPnP.Modernization.Framework.Entities;
using SharePointPnP.Modernization.Framework.Pages;
using SharePointPnP.Modernization.Framework.Telemetry;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace SharePointPnP.Modernization.Framework.Transform
{

    /// <summary>
    /// Transforms a classic wiki/webpart page into a modern client side page
    /// </summary>
    public class PageTransformator : BasePageTransformator
    {

        #region Construction
        /// <summary>
        /// Creates a page transformator instance with a target destination of a target web e.g. Modern/Communication Site
        /// </summary>
        /// <param name="sourceClientContext">ClientContext of the site holding the page</param>
        /// <param name="targetClientContext">ClientContext of the site that will receive the modernized page</param>
        public PageTransformator(ClientContext sourceClientContext, ClientContext targetClientContext) : this(sourceClientContext, targetClientContext, "webpartmapping.xml")
        {

        }

        /// <summary>
        /// Creates a page transformator instance
        /// </summary>
        /// <param name="sourceClientContext">ClientContext of the site holding the page</param>
        public PageTransformator(ClientContext sourceClientContext) : this(sourceClientContext, null, "webpartmapping.xml")
        {
        }

        /// <summary>
        /// Creates a page transformator instance
        /// </summary>
        /// <param name="sourceClientContext">ClientContext of the site holding the page</param>
        /// <param name="pageTransformationFile">Used page mapping file</param>
        public PageTransformator(ClientContext sourceClientContext, string pageTransformationFile) : this(sourceClientContext, null, pageTransformationFile)
        {

        }

        /// <summary>
        /// Creates a page transformator instance
        /// </summary>
        /// <param name="sourceClientContext">ClientContext of the site holding the page</param>
        /// <param name="targetClientContext">ClientContext of the site that will receive the modernized page</param>
        /// <param name="pageTransformationFile">Used page mapping file</param>
        public PageTransformator(ClientContext sourceClientContext, ClientContext targetClientContext, string pageTransformationFile)
        {

#if DEBUG && MEASURE && MEASURE
            InitMeasurement();
#endif

            this.sourceClientContext = sourceClientContext;
            this.targetClientContext = targetClientContext;

            this.version = GetVersion();
            this.pageTelemetry = new PageTelemetry(version);

            using (Stream schema = typeof(PageTransformator).Assembly.GetManifestResourceStream("SharePointPnP.Modernization.Framework.webpartmapping.xsd"))
            {
                // Load xml mapping data
                XmlSerializer xmlMapping = new XmlSerializer(typeof(PageTransformation));
                using (var stream = new FileStream(pageTransformationFile, FileMode.Open))
                {
                    // Ensure the provided file complies with the current schema
                    ValidateSchema(schema, stream);

                    // All good so it seems
                    this.pageTransformation = (PageTransformation)xmlMapping.Deserialize(stream);
                }
            }
        }

        /// <summary>
        /// Creates a page transformator instance
        /// </summary>
        /// <param name="sourceClientContext">ClientContext of the site holding the page</param>
        /// <param name="pageTransformationModel">Page transformation model</param>
        public PageTransformator(ClientContext sourceClientContext, PageTransformation pageTransformationModel) : this(sourceClientContext, null, pageTransformationModel)
        {

        }

        /// <summary>
        /// Creates a page transformator instance
        /// </summary>
        /// <param name="sourceClientContext">ClientContext of the site holding the page</param>
        /// <param name="targetClientContext">ClientContext of the site that will receive the modernized page</param>
        /// <param name="pageTransformationModel">Page transformation model</param>
        public PageTransformator(ClientContext sourceClientContext, ClientContext targetClientContext, PageTransformation pageTransformationModel)
        {

#if DEBUG && MEASURE
            InitMeasurement();
#endif

            this.sourceClientContext = sourceClientContext;
            this.targetClientContext = targetClientContext;

            this.version = GetVersion();
            this.pageTelemetry = new PageTelemetry(version);

            this.pageTransformation = pageTransformationModel;
        }
        #endregion

        /// <summary>
        /// Transform the page
        /// </summary>
        /// <param name="pageTransformationInformation">Information about the page to transform</param>
        /// <returns>The path to the created modern page</returns>
        public string Transform(PageTransformationInformation pageTransformationInformation)
        {
            SetPageId(Guid.NewGuid().ToString());

            var logsForSettings = pageTransformationInformation.DetailSettingsAsLogEntries();
            logsForSettings?.ForEach(o => Log(o, LogLevel.Information));

            #region Check for Target Site Context
            var hasTargetContext = targetClientContext != null;
            #endregion

            #region Input validation
            string pageType = null;
            if (pageTransformationInformation.SourceFile != null && pageTransformationInformation.SourcePage == null)
            {
                //TODO: extend check to ensure it's a real web part page
                isRootPage = IsRootPage(pageTransformationInformation.SourceFile);

                if (isRootPage)
                {
                    LogInfo(LogStrings.PageLivesOutsideOfALibrary, LogStrings.Heading_InputValidation);

                    // This always is a web part page
                    pageType = "WebPartPage";

                    // Item level permission copy makes no sense here
                    pageTransformationInformation.KeepPageSpecificPermissions = false;

                    // Same for swap pages, we don't support this as the pages live in a different location
                    pageTransformationInformation.TargetPageTakesSourcePageName = false;
                }
                else
                {
                    LogError(LogStrings.Error_BasicASPXPageCannotTransform, LogStrings.Heading_InputValidation);
                    throw new ArgumentException(LogStrings.Error_BasicASPXPageCannotTransform);
                }
            }
            else
            {
                if (pageTransformationInformation.SourcePage == null)
                {
                    LogError(LogStrings.Error_SourcePageNotFound, LogStrings.Heading_InputValidation);
                    throw new ArgumentNullException(LogStrings.Error_SourcePageNotFound);
                }

                // Validate page and it's eligibility for transformation
                if (!pageTransformationInformation.SourcePage.FieldExistsAndUsed(Constants.FileRefField) || !pageTransformationInformation.SourcePage.FieldExistsAndUsed(Constants.FileLeafRefField))
                {
                    LogError(LogStrings.Error_PageNotValidMissingFileRef, LogStrings.Heading_InputValidation);
                    throw new ArgumentException(LogStrings.Error_PageNotValidMissingFileRef);
                }

                pageType = pageTransformationInformation.SourcePage.PageType();

                if (pageType.Equals("ClientSidePage", StringComparison.InvariantCultureIgnoreCase))
                {
                    LogError(LogStrings.Error_SourcePageIsModern, LogStrings.Heading_InputValidation);
                    throw new ArgumentException(LogStrings.Error_SourcePageIsModern);
                }

                if (pageType.Equals("AspxPage", StringComparison.InvariantCultureIgnoreCase))
                {
                    LogError(LogStrings.Error_BasicASPXPageCannotTransform, LogStrings.Heading_InputValidation);
                    throw new ArgumentException(LogStrings.Error_BasicASPXPageCannotTransform);
                }

                if (pageType.Equals("PublishingPage", StringComparison.InvariantCultureIgnoreCase))
                {
                    LogError(LogStrings.Error_PublishingPagesNotYetSupported, LogStrings.Heading_InputValidation);
                    throw new ArgumentException(LogStrings.Error_PublishingPagesNotYetSupported);
                }
            }

            if (hasTargetContext)
            {
                // If we're transforming into another site collection the "revert to old page" model does not exist as the 
                // old page is not present in there. Also adding the page transformation banner does not make sense for the same reason
                if (pageTransformationInformation.ModernizationCenterInformation != null && pageTransformationInformation.ModernizationCenterInformation.AddPageAcceptBanner)
                {
                    LogError(LogStrings.Error_CannotUsePageAcceptBannerCrossSite, LogStrings.Heading_InputValidation);
                    throw new ArgumentException(LogStrings.Error_CannotUsePageAcceptBannerCrossSite);
                }
            }

            // Disable cross-farm item level permissions from copying
            CrossFarmTransformationValidation(pageTransformationInformation);

            LogDebug(LogStrings.ValidationChecksComplete, LogStrings.Heading_InputValidation);

            #endregion

            try
            {

                #region Telemetry
#if DEBUG && MEASURE
            Start();
#endif
                DateTime transformationStartDateTime = DateTime.Now;

                LogDebug(LogStrings.LoadingClientContextObjects, LogStrings.Heading_SharePointConnection);
                LoadClientObject(sourceClientContext, false);

                LogInfo($"{sourceClientContext.Web.GetUrl()}", LogStrings.Heading_Summary, LogEntrySignificance.SourceSiteUrl);

                if (hasTargetContext)
                {
                    LogDebug(LogStrings.LoadingTargetClientContext, LogStrings.Heading_SharePointConnection);
                    LoadClientObject(targetClientContext,true);

                    if (sourceClientContext.Site.Id.Equals(targetClientContext.Site.Id))
                    {
                        // Oops, seems source and target point to the same site collection...switch back the "source only" mode
                        targetClientContext = null;
                        hasTargetContext = false;
                        LogWarning(LogStrings.Error_FallBackToSameSiteTransfer, LogStrings.Heading_SharePointConnection);
                    }
                    else
                    {
                        // Ensure that the newly created page in the other site collection gets the same name as the source page
                        LogInfo(LogStrings.Error_OverridingTagePageTakesSourcePageName, LogStrings.Heading_SharePointConnection);
                        pageTransformationInformation.TargetPageTakesSourcePageName = true;
                    }

                    LogInfo($"{targetClientContext.Web.GetUrl()}", LogStrings.Heading_Summary, LogEntrySignificance.TargetSiteUrl);
                }

                // Need to add further validation for target template
                if (hasTargetContext &&
                   (targetClientContext.Web.WebTemplate != "SITEPAGEPUBLISHING" && targetClientContext.Web.WebTemplate != "STS" && targetClientContext.Web.WebTemplate != "GROUP"))
                {

                    LogError(LogStrings.Error_CrossSiteTransferTargetsNonModernSite);
                    throw new ArgumentException(LogStrings.Error_CrossSiteTransferTargetsNonModernSite, LogStrings.Heading_SharePointConnection);
                }

                LogInfo($"{GetFieldValue(pageTransformationInformation, Constants.FileRefField).ToLower()}", LogStrings.Heading_Summary, LogEntrySignificance.SourcePage);

#if DEBUG && MEASURE
            Stop("Telemetry");
#endif
                #endregion

                #region Page creation
                // Detect if the page is living inside a folder
                LogDebug(LogStrings.DetectIfPageIsInFolder, LogStrings.Heading_PageCreation);
                string pageFolder = "";

                if (FieldExistsAndIsUsed(pageTransformationInformation, Constants.FileDirRefField))
                {
                    var fileRefFieldValue = GetFieldValue(pageTransformationInformation, Constants.FileDirRefField);

                    if (fileRefFieldValue.ToLower().Contains("/sitepages"))
                    {
                        pageFolder = fileRefFieldValue.Replace($"{sourceClientContext.Web.ServerRelativeUrl.TrimEnd(new[] { '/' })}/SitePages", "").Trim();
                    }
                    else
                    {
                        // Page was living in another list, leave the list name as that will be the folder hosting the modern file in SitePages.
                        // This convention is used to avoid naming conflicts
                        pageFolder = fileRefFieldValue.Replace($"{sourceClientContext.Web.ServerRelativeUrl}", "").Trim();
                    }

                    if (pageFolder.Length > 0)
                    {
                        if (pageFolder.Contains("/"))
                        {
                            if (pageFolder == "/")
                            {
                                pageFolder = "";
                            }
                            else
                            {
                                pageFolder = pageFolder.Substring(1);
                            }
                        }

                        // Add a trailing slash
                        pageFolder = pageFolder + "/";

                        LogInfo(LogStrings.PageIsLocatedInFolder, LogStrings.Heading_PageCreation);
                    }

                    if (isRootPage)
                    {
                        pageFolder = "Root/";
                        pageTransformationInformation.TargetPageName = $"{GetFieldValue(pageTransformationInformation, Constants.FileLeafRefField)}";
                    }
                }
                pageTransformationInformation.Folder = pageFolder;

                // If no targetname specified then we'll come up with one
                if (string.IsNullOrEmpty(pageTransformationInformation.TargetPageName))
                {
                    if (string.IsNullOrEmpty(pageTransformationInformation.TargetPagePrefix))
                    {
                        LogInfo(LogStrings.NoTargetNameUsingDefaultPrefix, LogStrings.Heading_PageCreation);
                        pageTransformationInformation.SetDefaultTargetPagePrefix();
                    }

                    if (hasTargetContext)
                    {
                        LogInfo(LogStrings.CrossSiteInUseUsingOriginalFileName, LogStrings.Heading_PageCreation);
                        pageTransformationInformation.TargetPageName = $"{GetFieldValue(pageTransformationInformation, Constants.FileLeafRefField)}";
                    }
                    else
                    {
                        LogInfo(LogStrings.UsingSuppliedPrefix, LogStrings.Heading_PageCreation);

                        pageTransformationInformation.TargetPageName = $"{pageTransformationInformation.TargetPagePrefix}{GetFieldValue(pageTransformationInformation, Constants.FileLeafRefField)}";
                    }

                }

                // Check if page name is free to use
#if DEBUG && MEASURE
            Start();
#endif
                bool pageExists = false;
                ClientSidePage targetPage = null;
                List pagesLibrary = null;
                Microsoft.SharePoint.Client.File existingFile = null;

                //The determines of the target client context has been specified and use that to generate the target page
                var context = hasTargetContext ? targetClientContext : sourceClientContext;

                try
                {
                    LogDebug(LogStrings.LoadingExistingPageIfExists, LogStrings.Heading_PageCreation);

                    // Just try to load the page in the fastest possible manner, we only want to see if the page exists or not
                    existingFile = Load(sourceClientContext, pageTransformationInformation, out pagesLibrary, targetClientContext);
                    pageExists = true;
                }
                catch (Exception ex)
                {
                    if (ex is ArgumentException)
                    {
                        LogInfo(LogStrings.CheckPageExistsError, LogStrings.Heading_PageCreation);
                    }
                    else
                    {
                        LogError(LogStrings.CheckPageExistsError, LogStrings.Heading_PageCreation, ex, true);
                    }

                }
#if DEBUG && MEASURE
            Stop("Load Page");
#endif

                if (pageExists)
                {
                    LogInfo(LogStrings.PageAlreadyExistsInTargetLocation, LogStrings.Heading_PageCreation);

                    if (!pageTransformationInformation.Overwrite)
                    {
                        var message = $"{LogStrings.PageNotOverwriteIfExists}  {pageTransformationInformation.TargetPageName}.";
                        LogError(message, LogStrings.Heading_PageCreation);
                        throw new ArgumentException(message);
                    }
                }

                // Create the client side page

                targetPage = context.Web.AddClientSidePage($"{pageTransformationInformation.Folder}{pageTransformationInformation.TargetPageName}");
                LogInfo($"{LogStrings.ModernPageCreated} ", LogStrings.Heading_PageCreation);
                #endregion

                #region Home page handling
#if DEBUG && MEASURE
            Start();
#endif
                LogDebug(LogStrings.TransformCheckIfPageIsHomePage, LogStrings.Heading_HomePageHandling);

                bool replacedByOOBHomePage = false;
                // Check if the transformed page is the web's home page
                if (sourceClientContext.Web.RootFolder.IsPropertyAvailable("WelcomePage") && !string.IsNullOrEmpty(sourceClientContext.Web.RootFolder.WelcomePage))
                {
                    LogInfo(LogStrings.WelcomePageSettingsIsPresent, LogStrings.Heading_HomePageHandling);

                    var homePageUrl = sourceClientContext.Web.RootFolder.WelcomePage;
                    var homepageName = Path.GetFileName(sourceClientContext.Web.RootFolder.WelcomePage);
                    if (homepageName.Equals(GetFieldValue(pageTransformationInformation, Constants.FileLeafRefField), StringComparison.InvariantCultureIgnoreCase))
                    {
                        LogInfo(LogStrings.TransformSourcePageIsHomePage, LogStrings.Heading_HomePageHandling);

                        targetPage.LayoutType = ClientSidePageLayoutType.Home;
                        if (pageTransformationInformation.ReplaceHomePageWithDefaultHomePage)
                        {
                            targetPage.KeepDefaultWebParts = true;
                            replacedByOOBHomePage = true;

                            LogInfo(LogStrings.TransformSourcePageHomePageUsingStock,
                                LogStrings.Heading_HomePageHandling);
                        }
                    }
                    else
                    {
                        LogInfo(LogStrings.TransformSourcePageIsNotHomePage, LogStrings.Heading_HomePageHandling);
                    }
                }
#if DEBUG && MEASURE
            Stop(LogStrings.Heading_HomePageHandling);
#endif
                #endregion

                #region Article page handling

                if (!replacedByOOBHomePage)
                {
                    LogInfo(LogStrings.TransformSourcePageAsArticlePage, LogStrings.Heading_ArticlePageHandling);

                    #region Configure header from target page
#if DEBUG && MEASURE
                Start();
#endif
                    if (pageTransformationInformation.PageHeader == null || pageTransformationInformation.PageHeader.Type == ClientSidePageHeaderType.None)
                    {
                        LogInfo(LogStrings.TransformArticleSetHeaderToNone, LogStrings.Heading_ArticlePageHandling);

                        targetPage.RemovePageHeader();
                    }
                    else if (pageTransformationInformation.PageHeader.Type == ClientSidePageHeaderType.Default)
                    {
                        LogInfo(LogStrings.TransformArticleSetHeaderToDefault, LogStrings.Heading_ArticlePageHandling);

                        targetPage.SetDefaultPageHeader();
                    }
                    else if (pageTransformationInformation.PageHeader.Type == ClientSidePageHeaderType.Custom)
                    {
                        LogInfo($"{LogStrings.TransformArticleSetHeaderToCustom} " +
                                $"{LogStrings.TransformArticleHeaderImageUrl} {pageTransformationInformation.PageHeader.ImageServerRelativeUrl} ", LogStrings.Heading_ArticlePageHandling);

                        targetPage.SetCustomPageHeader(pageTransformationInformation.PageHeader.ImageServerRelativeUrl, pageTransformationInformation.PageHeader.TranslateX, pageTransformationInformation.PageHeader.TranslateY);
                    }
#if DEBUG && MEASURE
                Stop("Target page header");
#endif
                    #endregion

                    #region Analysis of the source page
#if DEBUG && MEASURE
                Start();
#endif
                    // Analyze the source page
                    Tuple<PageLayout, List<WebPartEntity>> pageData = null;

                    if (pageType.Equals("WikiPage", StringComparison.InvariantCultureIgnoreCase))
                    {
                        LogInfo($"{LogStrings.TransformSourcePageIsWikiPage} - {LogStrings.TransformSourcePageAnalysing}", LogStrings.Heading_ArticlePageHandling);

                        pageData = new WikiPage(pageTransformationInformation.SourcePage, pageTransformation).Analyze();

                        // Wiki pages can contain embedded images and videos, which is not supported by the target RTE...split wiki text blocks so the transformator can handle the images and videos as separate web parts
                        LogInfo(LogStrings.WikiTextContainsImagesVideosReferences, LogStrings.Heading_ArticlePageHandling);
                    }
                    else if (pageType.Equals("WebPartPage", StringComparison.InvariantCultureIgnoreCase))
                    {
                        LogInfo($"{LogStrings.TransformSourcePageIsWebPartPage} {LogStrings.TransformSourcePageAnalysing}", LogStrings.Heading_ArticlePageHandling);

                        var spVersion = GetVersion(sourceClientContext);

                        if (spVersion == SPVersion.SP2010 || spVersion == SPVersion.SP2013Legacy || spVersion == SPVersion.SP2016Legacy)
                        {
                            pageData = new WebPartPageOnPremises(pageTransformationInformation.SourcePage, pageTransformationInformation.SourceFile, pageTransformation).Analyze(true);
                        }
                        else
                        {
                            pageData = new WebPartPage(pageTransformationInformation.SourcePage, pageTransformationInformation.SourceFile, pageTransformation).Analyze(true);
                        }
                    }

                    // Analyze the "text" parts (wikitext parts and text in content editor web parts)
                    pageData = new Tuple<PageLayout, List<WebPartEntity>>(pageData.Item1, new WikiHtmlTransformator(this.sourceClientContext, targetPage, pageTransformationInformation as BaseTransformationInformation, base.RegisteredLogObservers).TransformPlusSplit(pageData.Item2, pageTransformationInformation.HandleWikiImagesAndVideos, pageTransformationInformation.AddTableListImageAsImageWebPart));

#if DEBUG && MEASURE
                Stop("Analyze page");
#endif
                    #endregion

                    #region Page title configuration
#if DEBUG && MEASURE
                Start();
#endif
                    // Set page title
                    if (pageType.Equals("WikiPage", StringComparison.InvariantCultureIgnoreCase))
                    {
                        SetPageTitle(pageTransformationInformation, targetPage);
                    }
                    else if (pageType.Equals("WebPartPage"))
                    {
                        bool titleFound = false;
                        var titleBarWebPart = pageData.Item2.Where(p => p.Type == WebParts.TitleBar).FirstOrDefault();
                        if (titleBarWebPart != null)
                        {
                            if (titleBarWebPart.Properties.ContainsKey("HeaderTitle") && !string.IsNullOrEmpty(titleBarWebPart.Properties["HeaderTitle"]))
                            {
                                var title = titleBarWebPart.Properties["HeaderTitle"];

                                LogInfo($"{LogStrings.TransformPageModernTitle} {title}", LogStrings.Heading_ArticlePageHandling);
                                targetPage.PageTitle = title;
                                titleFound = true;
                            }
                        }

                        if (!titleFound)
                        {
                            SetPageTitle(pageTransformationInformation, targetPage);
                        }
                    }

                    if (pageTransformationInformation.PageTitleOverride != null)
                    {
                        var title = pageTransformationInformation.PageTitleOverride(targetPage.PageTitle);
                        targetPage.PageTitle = title;

                        LogInfo($"{LogStrings.TransformPageTitleOverride} - page title: {title}", LogStrings.Heading_ArticlePageHandling);
                    }
#if DEBUG && MEASURE
                Stop("Set page title");
#endif
                    #endregion

                    #region Page layout configuration
#if DEBUG && MEASURE
                Start();
#endif
                    // Use the default layout transformator
                    ILayoutTransformator layoutTransformator = new LayoutTransformator(targetPage);

                    // Do we have an override?
                    if (pageTransformationInformation.LayoutTransformatorOverride != null)
                    {
                        LogInfo(LogStrings.TransformLayoutTransformatorOverride, LogStrings.Heading_ArticlePageHandling);
                        layoutTransformator = pageTransformationInformation.LayoutTransformatorOverride(targetPage);
                    }

                    // Apply the layout to the page
                    layoutTransformator.Transform(pageData);
#if DEBUG && MEASURE
                Stop("Page layout");
#endif
                    #endregion

                    #region Page Banner creation
                    if (!pageTransformationInformation.TargetPageTakesSourcePageName)
                    {

                        if (pageTransformationInformation.ModernizationCenterInformation != null && pageTransformationInformation.ModernizationCenterInformation.AddPageAcceptBanner)
                        {

#if DEBUG && MEASURE
                        Start();
#endif

                            // Bump the row values for the existing web parts as we've inserted a new section
                            foreach (var section in targetPage.Sections)
                            {
                                section.Order = section.Order + 1;
                            }

                            // Add new section for banner part
                            targetPage.Sections.Insert(0, new CanvasSection(targetPage, CanvasSectionTemplate.OneColumn, 0));

                            // Bump the row values for the existing web parts as we've inserted a new section
                            foreach (var webpart in pageData.Item2.Where(c => !c.IsClosed))
                            {
                                webpart.Row = webpart.Row + 1;
                            }


                            var sourcePageUrl = GetFieldValue(pageTransformationInformation, Constants.FileRefField);
                            var orginalSourcePageName = GetFieldValue(pageTransformationInformation, Constants.FileLeafRefField);
                            Uri host = new Uri(sourceClientContext.Web.GetUrl());

                            string path = $"{host.Scheme}://{host.DnsSafeHost}{sourcePageUrl.Replace(GetFieldValue(pageTransformationInformation, Constants.FileLeafRefField), "")}";

                            // Add "fake" banner web part that then will be transformed onto the page
                            Dictionary<string, string> props = new Dictionary<string, string>(2)
                        {
                            { "SourcePage", $"{path}{orginalSourcePageName}" },
                            { "TargetPage", $"{path}{pageTransformationInformation.TargetPageName}" }
                        };

                            WebPartEntity bannerWebPart = new WebPartEntity()
                            {
                                Type = WebParts.PageAcceptanceBanner,
                                Column = 1,
                                Row = 1,
                                Title = "",
                                Order = 0,
                                Properties = props,
                            };
                            pageData.Item2.Insert(0, bannerWebPart);
                            LogInfo(LogStrings.TransformAddedPageAcceptBanner, LogStrings.Heading_ArticlePageHandling);

#if DEBUG && MEASURE
                        Stop("Page Banner");
#endif
                        }
                    }
                    #endregion

                    #region Content transformation

                    LogDebug(LogStrings.PreparingContentTransformation, LogStrings.Heading_ArticlePageHandling);

#if DEBUG && MEASURE
                Start();
#endif
                    // Use the default content transformator
                    IContentTransformator contentTransformator = new ContentTransformator(sourceClientContext, targetPage, pageTransformation, pageTransformationInformation as BaseTransformationInformation, base.RegisteredLogObservers);

                    // Do we have an override?
                    if (pageTransformationInformation.ContentTransformatorOverride != null)
                    {
                        LogInfo(LogStrings.TransformUsingContentTransformerOverride, LogStrings.Heading_ArticlePageHandling);

                        contentTransformator = pageTransformationInformation.ContentTransformatorOverride(targetPage, pageTransformation);
                    }

                    LogInfo(LogStrings.TransformingContentStart, LogStrings.Heading_ArticlePageHandling);

                    // Run the content transformator
                    contentTransformator.Transform(pageData.Item2.Where(c => !c.IsClosed).ToList());

                    LogInfo(LogStrings.TransformingContentEnd, LogStrings.Heading_ArticlePageHandling);
#if DEBUG && MEASURE
                Stop("Content transformation");
#endif
                    #endregion

                    #region Text/Section/Column cleanup
                    // Drop "empty" text parts. Wiki pages tend to have a lot of text parts just containing div's and BR's...no point in keep those as they generate to much whitespace
                    RemoveEmptyTextParts(targetPage);

                    // Remove empty sections and columns to optimize screen real estate
                    if (pageTransformationInformation.RemoveEmptySectionsAndColumns)
                    {
                        RemoveEmptySectionsAndColumns(targetPage);
                    }
                    #endregion
                }
                #endregion

                #region Page persisting + permissions
                #region Save the page
#if DEBUG && MEASURE
            Start();
#endif
                // Persist the client side page
                if (hasTargetContext)
                {
                    var pageName = $"{pageTransformationInformation.Folder}{pageTransformationInformation.TargetPageName}";

                    targetPage.Save(pageName);

                    LogInfo($"{LogStrings.TransformSavedPageInCrossSiteCollection}: {pageName}", LogStrings.Heading_ArticlePageHandling);
                }
                else
                {
                    var pageName = $"{pageTransformationInformation.Folder}{pageTransformationInformation.TargetPageName}";

                    targetPage.Save(pageName, existingFile, pagesLibrary);

                    LogInfo($"{LogStrings.TransformSavedPage}: {pageName}", LogStrings.Heading_ArticlePageHandling);
                }

#if DEBUG && MEASURE
            Stop("Persist page");
#endif
                #endregion

                #region Page metadata handling
                // Temporary removal of metadata copy for cross site.
                if (pageTransformationInformation.CopyPageMetadata && !hasTargetContext)
                {
#if DEBUG && MEASURE
                Start();
#endif
                    // Copy the page metadata 
                    CopyPageMetadata(pageTransformationInformation, targetPage, pagesLibrary);
#if DEBUG && MEASURE
                Stop("Page metadata handling");
#endif
                }
                #endregion

                #region Permission handling
                ListItemPermission listItemPermissionsToKeep = null;
                if (pageTransformationInformation.KeepPageSpecificPermissions)
                {
#if DEBUG && MEASURE
                Start();
#endif
                    // Check if we do have item level permissions we want to take over
                    listItemPermissionsToKeep = GetItemLevelPermissions(hasTargetContext, pagesLibrary, pageTransformationInformation.SourcePage, targetPage.PageListItem);

                    if (!pageTransformationInformation.TargetPageTakesSourcePageName || hasTargetContext)
                    {
                        // If we're not doing a page name swap now we need to update the target item with the needed item level permissions.                    
                        // When creating the page in another site collection we'll always want to copy item level permissions if specified
                        ApplyItemLevelPermissions(hasTargetContext, targetPage.PageListItem, listItemPermissionsToKeep);
                    }
#if DEBUG && MEASURE
                Stop("Permission handling");
#endif
                }
                #endregion

                #region Page Publishing
                // Tag the file with a page modernization version stamp
                string serverRelativePathForModernPage = ReturnModernPageServerRelativeUrl(pageTransformationInformation, hasTargetContext);
                try
                {
                    var targetPageFile = context.Web.GetFileByServerRelativeUrl(serverRelativePathForModernPage);
                    context.Load(targetPageFile, p => p.Properties);
                    targetPageFile.Properties["sharepointpnp_pagemodernization"] = this.version;
                    targetPageFile.Update();

                    if (pageTransformationInformation.PublishCreatedPage)
                    {
                        // Try to publish, if publish is not needed then this will return an error that we'll be ignoring
                        targetPageFile.Publish("Page modernization initial publish");
                    }

                    // Send both the property update and publish as a single operation to SharePoint
                    context.ExecuteQueryRetry();
                }
                catch (Exception ex)
                {
                    // Eat exceptions as this is not critical for the generated page
                    LogWarning(LogStrings.Warning_NonCriticalErrorDuringVersionStampAndPublish, LogStrings.Heading_ArticlePageHandling);
                }

                // Disable page comments on the create page, if needed
                if (pageTransformationInformation.DisablePageComments)
                {
                    targetPage.DisableComments();
                    LogInfo(LogStrings.TransformDisablePageComments, LogStrings.Heading_ArticlePageHandling);
                }

                #endregion

                #region Page name switching
                // All went well so far...swap pages if that's needed. When copying to another site collection this step is not needed
                // as the created page already has the final name
                if (pageTransformationInformation.TargetPageTakesSourcePageName && !hasTargetContext)
                {
#if DEBUG && MEASURE
                Start();
#endif
                    //Load the source page
                    SwapPages(pageTransformationInformation, listItemPermissionsToKeep);
#if DEBUG && MEASURE
                Stop("Pagename swap");
#endif
                }
                #endregion

                #region Telemetry
                if (!pageTransformationInformation.SkipTelemetry && this.pageTelemetry != null)
                {
                    TimeSpan duration = DateTime.Now.Subtract(transformationStartDateTime);
                    this.pageTelemetry.LogTransformationDone(duration, pageType, pageTransformationInformation);
                    this.pageTelemetry.Flush();
                }

                LogInfo(LogStrings.TransformComplete, LogStrings.Heading_PageCreation);
                #endregion

                #region Closing
                CacheManager.Instance.SetLastUsedTransformator(this);
                return serverRelativePathForModernPage;
                #endregion

                #endregion
            }
            catch (Exception ex)
            {
                LogError(LogStrings.CriticalError_ErrorOccurred, LogStrings.Heading_Summary, ex, isCriticalException: true);

                // Throw exception if there's no registered log observers
                if (base.RegisteredLogObservers.Count == 0)
                {
                    throw;
                }
            }

            return string.Empty;
        }

        /// <summary>
        /// Performs the logic needed to swap a genered Migrated_Page.aspx to Page.aspx and then Page.aspx to Old_Page.aspx
        /// </summary>
        /// <param name="pageTransformationInformation">Information about the page to transform</param>
        public void SwapPages(PageTransformationInformation pageTransformationInformation, ListItemPermission listItemPermissionsToKeep)
        {
            LogInfo("Swapping pages", LogStrings.Heading_SwappingPages);
            var sourcePageUrl = GetFieldValue(pageTransformationInformation, Constants.FileRefField);
            var orginalSourcePageName = GetFieldValue(pageTransformationInformation, Constants.FileLeafRefField);
            
            string sourcePath = sourcePageUrl.Replace(GetFieldValue(pageTransformationInformation, Constants.FileLeafRefField), "");
            string targetPath = sourcePath;

            if (!sourcePath.ToLower().Contains("/sitepages"))
            {
                // Source file was living outside of the site pages library
                targetPath = sourcePath.Replace(sourceClientContext.Web.ServerRelativeUrl, "");
                targetPath = $"{sourceClientContext.Web.ServerRelativeUrl}/SitePages{targetPath}";
            }

            var sourcePage = this.sourceClientContext.Web.GetFileByServerRelativeUrl(sourcePageUrl);
            this.sourceClientContext.Load(sourcePage);
            this.sourceClientContext.ExecuteQueryRetry();

            if (string.IsNullOrEmpty(pageTransformationInformation.SourcePagePrefix))
            {
                LogInfo("Using default source page prefix", LogStrings.Heading_SwappingPages);
                pageTransformationInformation.SetDefaultSourcePagePrefix();
            }
            var newSourcePageUrl = $"{pageTransformationInformation.SourcePagePrefix}{GetFieldValue(pageTransformationInformation, Constants.FileLeafRefField)}";


            // Rename source page using the sourcepageprefix
            // STEP1: First copy the source page to a new name. We on purpose use CopyTo as we want to avoid that "linked" url's get 
            //        patched up during a MoveTo operation as that would also patch the url's in our new modern page
            var step1Path = $"{sourcePath}{newSourcePageUrl}";
            sourcePage.CopyTo(step1Path, true);
            this.sourceClientContext.ExecuteQueryRetry();
            LogInfo($"{LogStrings.TransformSwappingPageStep1}: {step1Path}", LogStrings.Heading_SwappingPages);

            // Restore the item level permissions on the copied page (if any)
            if (pageTransformationInformation.KeepPageSpecificPermissions && listItemPermissionsToKeep != null)
            {
                LogInfo(LogStrings.TransformSwappingPageRestorePermissions, LogStrings.Heading_SwappingPages);

                // load the copied target file
                var newSource = this.sourceClientContext.Web.GetFileByServerRelativeUrl($"{sourcePath}{newSourcePageUrl}");
                this.sourceClientContext.Load(newSource);
                this.sourceClientContext.Load(newSource.ListItemAllFields, p => p.RoleAssignments);
                this.sourceClientContext.ExecuteQueryRetry();

                // Reload source page
                ApplyItemLevelPermissions(false, newSource.ListItemAllFields, listItemPermissionsToKeep, alwaysBreakItemLevelPermissions: true);
            }

            //Load the created target page
            var targetPageUrl = $"{targetPath}{pageTransformationInformation.TargetPageName}";
            var targetPageFile = this.sourceClientContext.Web.GetFileByServerRelativeUrl(targetPageUrl);
            this.sourceClientContext.Load(targetPageFile);
            this.sourceClientContext.ExecuteQueryRetry();

            LogInfo(LogStrings.TransformSwappingPageStep2, LogStrings.Heading_SwappingPages);

            // STEP2: Fix possible navigation entries to point to the "copied" source page first
            // Rename the target page to the original source page name
            // CopyTo and MoveTo with option to overwrite first internally delete the file to overwrite, which
            // results in all page navigation nodes pointing to this file to be deleted. Hence let's point these
            // navigation entries first to the copied version of the page we just created
            this.sourceClientContext.Web.Context.Load(this.sourceClientContext.Web, w => w.Navigation.QuickLaunch, w => w.Navigation.TopNavigationBar);
            this.sourceClientContext.Web.Context.ExecuteQueryRetry();

            bool navWasFixed = false;
            IQueryable<NavigationNode> currentNavNodes = null;
            IQueryable<NavigationNode> globalNavNodes = null;
            var currentNavigation = this.sourceClientContext.Web.Navigation.QuickLaunch;
            var globalNavigation = this.sourceClientContext.Web.Navigation.TopNavigationBar;
            // Check for nav nodes
            currentNavNodes = currentNavigation.Where(n => n.Url.Equals(sourcePageUrl, StringComparison.InvariantCultureIgnoreCase));
            globalNavNodes = globalNavigation.Where(n => n.Url.Equals(sourcePageUrl, StringComparison.InvariantCultureIgnoreCase));

            if (currentNavNodes.Count() > 0 || globalNavNodes.Count() > 0)
            {
                navWasFixed = true;
                foreach (var node in currentNavNodes)
                {
                    node.Url = $"{sourcePath}{newSourcePageUrl}";
                    node.Update();
                }
                foreach (var node in globalNavNodes)
                {
                    node.Url = $"{sourcePath}{newSourcePageUrl}";
                    node.Update();
                }
                this.sourceClientContext.ExecuteQueryRetry();
                LogInfo(LogStrings.TransformSwappingPageUpdateNavigation, LogStrings.Heading_SwappingPages);
            }

            LogInfo(LogStrings.TransformSwappingPageStep3, LogStrings.Heading_SwappingPages);

            // STEP3: Now copy the created modern page over the original source page, at this point the new page has the same name as the original page had before transformation
            var step3Path = $"{targetPath}{orginalSourcePageName}";
            targetPageFile.CopyTo(step3Path, true);
            this.sourceClientContext.ExecuteQueryRetry();
            LogInfo($"{LogStrings.TransformSwappingPageStep3Path} :{step3Path}", LogStrings.Heading_SwappingPages);

            // Apply the item level permissions on the final page (if any)
            if (pageTransformationInformation.KeepPageSpecificPermissions && listItemPermissionsToKeep != null)
            {
                LogInfo(LogStrings.TransformSwappingPagesApplyItemPermissions, LogStrings.Heading_SwappingPages);

                // load the copied target file
                var newTarget = this.sourceClientContext.Web.GetFileByServerRelativeUrl($"{targetPath}{orginalSourcePageName}");
                this.sourceClientContext.Load(newTarget);
                this.sourceClientContext.Load(newTarget.ListItemAllFields, p => p.RoleAssignments);
                this.sourceClientContext.ExecuteQueryRetry();

                ApplyItemLevelPermissions(false, newTarget.ListItemAllFields, listItemPermissionsToKeep, alwaysBreakItemLevelPermissions: true);
            }

            // STEP4: Finish with restoring the page navigation: update the navlinks to point back the original page name
            LogInfo(LogStrings.TransformSwappingPagesStep4, LogStrings.Heading_SwappingPages);

            if (navWasFixed)
            {

                // Reload the navigation entries as did update them
                this.sourceClientContext.Web.Context.Load(this.sourceClientContext.Web, w => w.Navigation.QuickLaunch, w => w.Navigation.TopNavigationBar);
                this.sourceClientContext.Web.Context.ExecuteQueryRetry();

                currentNavigation = this.sourceClientContext.Web.Navigation.QuickLaunch;
                globalNavigation = this.sourceClientContext.Web.Navigation.TopNavigationBar;
                if (!string.IsNullOrEmpty($"{sourcePath}{newSourcePageUrl}"))
                {
                    currentNavNodes = currentNavigation.Where(n => n.Url.Equals($"{sourcePath}{newSourcePageUrl}", StringComparison.InvariantCultureIgnoreCase));
                    globalNavNodes = globalNavigation.Where(n => n.Url.Equals($"{sourcePath}{newSourcePageUrl}", StringComparison.InvariantCultureIgnoreCase));
                }

                foreach (var node in currentNavNodes)
                {
                    node.Url = sourcePageUrl;
                    node.Update();
                }
                foreach (var node in globalNavNodes)
                {
                    node.Url = sourcePageUrl;
                    node.Update();
                }
                this.sourceClientContext.ExecuteQueryRetry();
            }

            //STEP5: Conclude with deleting the originally created modern page as we did copy that already in step 3
            LogInfo(LogStrings.TransformSwappingPagesStep5, LogStrings.Heading_SwappingPages);
            targetPageFile.DeleteObject();
            this.sourceClientContext.ExecuteQueryRetry();

            //STEP6: if the source page lived outside of the site pages library then we also need to delete the original page from that spot
            if (sourcePath != targetPath)
            {
                LogInfo(LogStrings.TransformSwappingPagesStep6, LogStrings.Heading_SwappingPages);
                sourcePage.DeleteObject();
                this.sourceClientContext.ExecuteQueryRetry();
            }
        }

        /// <summary>
        /// Loads a page transformation model from file
        /// </summary>
        /// <param name="pageTransformationFile">File holding the page transformation model</param>
        /// <returns>Page transformation model</returns>
        public static PageTransformation LoadPageTransformationModel(string pageTransformationFile)
        {
            // Load xml mapping data
            XmlSerializer xmlMapping = new XmlSerializer(typeof(PageTransformation));
            using (var stream = new FileStream(pageTransformationFile, FileMode.Open))
            {
                return (PageTransformation)xmlMapping.Deserialize(stream);
            }
        }

        #region Helper methods
        private string ReturnModernPageServerRelativeUrl(PageTransformationInformation pageTransformationInformation, bool hasTargetContext)
        {
            string returnUrl = null;

            string originalSourcePageName = GetFieldValue(pageTransformationInformation, Constants.FileLeafRefField).ToLower();
            string sourcePath = GetFieldValue(pageTransformationInformation, Constants.FileRefField).ToLower().Replace(originalSourcePageName, "");
            string targetPath = sourcePath;

            if (hasTargetContext)
            {
                // Cross site collection transfer, new page always takes the name of the old page
                if (!sourcePath.Contains("/sitepages"))
                {
                    // Source file was living outside of the site pages library
                    targetPath = sourcePath.Replace(sourceClientContext.Web.ServerRelativeUrl.ToLower(), "");

                    if (pageTransformationInformation.SourceFile != null && pageTransformationInformation.SourcePage == null)
                    {
                        targetPath = targetPath + "root/";
                    }

                    targetPath = $"{targetClientContext.Web.ServerRelativeUrl.ToLower()}/sitepages{targetPath}";
                }
                else
                {
                    // Page was living inside the sitepages library
                    targetPath = sourcePath.Replace($"{sourceClientContext.Web.ServerRelativeUrl}/sitepages".ToLower(), "");
                    targetPath = $"{targetClientContext.Web.ServerRelativeUrl.ToLower()}/sitepages{targetPath}";
                }

                returnUrl = $"{targetPath}{originalSourcePageName}";
            }
            else
            {
                // In-place modernization
                if (!sourcePath.Contains("/sitepages"))
                {
                    // Source file was living outside of the site pages library
                    targetPath = sourcePath.Replace(sourceClientContext.Web.ServerRelativeUrl.ToLower(), "");

                    if (pageTransformationInformation.SourceFile != null && pageTransformationInformation.SourcePage == null)
                    {
                        targetPath = targetPath + "root/";
                    }

                    targetPath = $"{sourceClientContext.Web.ServerRelativeUrl}/sitepages{targetPath}".ToLower();
                }

                if (!pageTransformationInformation.TargetPageTakesSourcePageName)
                {
                    // New page uses a different name (e.g. Migrated_xxx.aspx)
                    returnUrl = $"{targetPath}{pageTransformationInformation.TargetPageName}".ToLower();
                }
                else
                {
                    // New page takes the name of the old page
                    returnUrl = $"{targetPath}{GetFieldValue(pageTransformationInformation, Constants.FileLeafRefField)}".ToLower();
                }
            }

            LogInfo($"{returnUrl}", LogStrings.Heading_Summary, LogEntrySignificance.TargetPage);
            return returnUrl;
        }

        private void SetPageTitle(PageTransformationInformation pageTransformationInformation, ClientSidePage targetPage)
        {
            if (FieldExistsAndIsUsed(pageTransformationInformation, Constants.FileLeafRefField))
            {
                string pageTitle = Path.GetFileNameWithoutExtension((GetFieldValue(pageTransformationInformation, Constants.FileLeafRefField)));
                if (!string.IsNullOrEmpty(pageTitle))
                {
                    pageTitle = pageTitle.First().ToString().ToUpper() + pageTitle.Substring(1);
                    targetPage.PageTitle = pageTitle;
                    LogInfo($"{LogStrings.TransformPageModernTitle} {pageTitle}", LogStrings.Heading_SetPageTitle);
                }
            }
        }

        private Microsoft.SharePoint.Client.File Load(ClientContext sourceContext, PageTransformationInformation pageTransformationInformation, out List pagesLibrary, ClientContext targetContext = null)
        {
            sourceContext.Web.EnsureProperty(w => w.ServerRelativeUrl);

            // Load the pages library and page file (if exists) in one go 
            if (GetVersion(sourceClientContext) == SPVersion.SP2010)
            {
                pagesLibrary = sourceContext.Web.GetSitePagesLibrary();
            }
            else
            {
                var listServerRelativeUrl = UrlUtility.Combine(sourceContext.Web.ServerRelativeUrl, "SitePages");
                pagesLibrary = sourceContext.Web.GetList(listServerRelativeUrl);
            }

            if (pageTransformationInformation.CopyPageMetadata)
            {
                sourceContext.Web.Context.Load(pagesLibrary, l => l.DefaultViewUrl, l => l.Id, l => l.BaseTemplate, l => l.OnQuickLaunch, l => l.DefaultViewUrl, l => l.Title, 
                                                  l => l.Hidden, l => l.EffectiveBasePermissions, l => l.RootFolder, l => l.RootFolder.ServerRelativeUrl, 
                                                  l => l.Fields.IncludeWithDefaultProperties(f => f.Id, f => f.Title, f => f.Hidden, f => f.InternalName, f => f.DefaultValue, f => f.Required));
            }
            else
            {
                sourceContext.Web.Context.Load(pagesLibrary, l => l.DefaultViewUrl, l => l.Id, l => l.BaseTemplate, l => l.OnQuickLaunch, l => l.DefaultViewUrl, l => l.Title, 
                                                  l => l.Hidden, l => l.EffectiveBasePermissions, l => l.RootFolder, l => l.RootFolder.ServerRelativeUrl);
            }

            var contextForFile = targetClientContext == null ? sourceClientContext : targetClientContext;
            var sitePagesServerRelativeUrl = UrlUtility.Combine(contextForFile.Web.ServerRelativeUrl, "sitepages");

            var file = contextForFile.Web.GetFileByServerRelativeUrl($"{sitePagesServerRelativeUrl}/{pageTransformationInformation.Folder}{pageTransformationInformation.TargetPageName}");
            contextForFile.Web.Context.Load(file, f => f.Exists, f => f.ListItemAllFields);
            contextForFile.ExecuteQueryRetry();

            if (pageTransformationInformation.KeepPageSpecificPermissions)
            {
                sourceContext.Load(pageTransformationInformation.SourcePage, p => p.HasUniqueRoleAssignments);
            }

            try
            {
                sourceContext.ExecuteQueryRetry();
            }
            catch (ServerException se)
            {
                if (se.ServerErrorTypeName == "System.IO.FileNotFoundException")
                {
                    pagesLibrary = null;
                }
                else
                {
                    throw;
                }
            }

            if (pagesLibrary == null)
            {
                LogError(LogStrings.Error_MissingSitePagesLibrary, LogStrings.Heading_Load);
                throw new ArgumentException(LogStrings.Error_MissingSitePagesLibrary);
            }

            if (!file.Exists)
            {
                LogInfo(LogStrings.TransformPageDoesNotExistInWeb, LogStrings.Heading_Load);
                throw new ArgumentException($"{pageTransformationInformation.TargetPageName} - {LogStrings.TransformPageDoesNotExistInWeb}");
            }

            return file;
        }

        private void ValidateSchema(Stream schema, FileStream stream)
        {
            // Load the template into an XDocument
            XDocument xml = XDocument.Load(stream);

            // Prepare the XML Schema Set
            XmlSchemaSet schemas = new XmlSchemaSet();
            schema.Seek(0, SeekOrigin.Begin);
            schemas.Add(Constants.PageTransformationSchema, new XmlTextReader(schema));

            // Set stream back to start
            stream.Seek(0, SeekOrigin.Begin);

            xml.Validate(schemas, (o, e) =>
            {
                LogError(string.Format(LogStrings.Error_WebPartMappingSchemaValidation, e.Message), LogStrings.Heading_PageTransformationInfomation, e.Exception);
                throw new Exception(string.Format(LogStrings.Error_MappingFileSchemaValidation, e.Message));
            });
        }
        #endregion

    }
}
