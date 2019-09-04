using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Pages;
using OfficeDevPnP.Core.Utilities;
using SharePointPnP.Modernization.Framework.Cache;
using SharePointPnP.Modernization.Framework.Entities;
using SharePointPnP.Modernization.Framework.Pages;
using SharePointPnP.Modernization.Framework.Telemetry;
using SharePointPnP.Modernization.Framework.Transform;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;

namespace SharePointPnP.Modernization.Framework.Publishing
{
    /// <summary>
    /// Transforms a classic publishing page into a modern client side page
    /// </summary>
    public class PublishingPageTransformator : BasePageTransformator
    {
        private PublishingPageTransformation publishingPageTransformation;
        private PageLayoutManager pageLayoutManager;
        private string publishingPagesLibraryName = null;

        #region Construction
        public PublishingPageTransformator(ClientContext sourceClientContext, ClientContext targetClientContext) : this(sourceClientContext, targetClientContext, "webpartmapping.xml", null)
        {
        }

        public PublishingPageTransformator(ClientContext sourceClientContext, ClientContext targetClientContext, string publishingPageTransformationFile) : this(sourceClientContext, targetClientContext, "webpartmapping.xml", publishingPageTransformationFile)
        {
        }

        public PublishingPageTransformator(ClientContext sourceClientContext, ClientContext targetClientContext, string pageTransformationFile, string publishingPageTransformationFile)
        {
#if DEBUG && MEASURE
            InitMeasurement();
#endif

            this.sourceClientContext = sourceClientContext ?? throw new ArgumentException("sourceClientContext must be provided.");
            this.targetClientContext = targetClientContext ?? throw new ArgumentException("targetClientContext must be provided."); ;

            this.version = GetVersion();
            this.pageTelemetry = new PageTelemetry(version);
            this.pageLayoutManager = new PageLayoutManager(base.RegisteredLogObservers);

            // Load xml mapping data
            XmlSerializer xmlMapping = new XmlSerializer(typeof(PageTransformation));
            using (var stream = new FileStream(pageTransformationFile, FileMode.Open))
            {
                this.pageTransformation = (PageTransformation)xmlMapping.Deserialize(stream);
            }

            // Load the page layout mapping data            
            this.publishingPageTransformation = this.pageLayoutManager.LoadDefaultPageLayoutMappingFile();

            if (!string.IsNullOrEmpty(publishingPageTransformationFile))
            {
                // Load page layout mapping from file
                var customPublishingPageTransformation = this.pageLayoutManager.LoadPageLayoutMappingFile(publishingPageTransformationFile);

                // Merge these custom layout mappings with the default ones
                this.publishingPageTransformation = this.pageLayoutManager.MergePageLayoutMappingFiles(this.publishingPageTransformation, customPublishingPageTransformation);
            }
        }

        public PublishingPageTransformator(ClientContext sourceClientContext, ClientContext targetClientContext, PageTransformation pageTransformationModel, PublishingPageTransformation customPublishingPageTransformationModel)
        {
#if DEBUG && MEASURE
            InitMeasurement();
#endif

            this.sourceClientContext = sourceClientContext ?? throw new ArgumentException("sourceClientContext must be provided.");
            this.targetClientContext = targetClientContext ?? throw new ArgumentException("targetClientContext must be provided."); ;

            this.version = GetVersion();
            this.pageTelemetry = new PageTelemetry(version);
            this.pageLayoutManager = new PageLayoutManager(base.RegisteredLogObservers);

            this.pageTransformation = pageTransformationModel;

            // Load the page layout mapping data            
            this.publishingPageTransformation = this.pageLayoutManager.LoadDefaultPageLayoutMappingFile();

            // Merge these custom layout mappings with the default ones
            if (customPublishingPageTransformationModel != null)
            {
                this.publishingPageTransformation = this.pageLayoutManager.MergePageLayoutMappingFiles(this.publishingPageTransformation, customPublishingPageTransformationModel);
            }
        }
        #endregion

        /// <summary>
        /// Transform the publishing page
        /// </summary>
        /// <param name="publishingPageTransformationInformation">Information about the publishing page to transform</param>
        /// <returns>The path to the created modern page</returns>
        public string Transform(PublishingPageTransformationInformation publishingPageTransformationInformation)
        {
            SetPageId(Guid.NewGuid().ToString());

            var logsForSettings = this.DetailSettingsAsLogEntries(publishingPageTransformationInformation);
            logsForSettings?.ForEach(o => Log(o, LogLevel.Information));

            #region Input validation
            if (publishingPageTransformationInformation.SourcePage == null)
            {
                LogError(LogStrings.Error_SourcePageNotFound, LogStrings.Heading_InputValidation);
                throw new ArgumentNullException(LogStrings.Error_SourcePageNotFound);
            }

            // Validate page and it's eligibility for transformation
            if (!publishingPageTransformationInformation.SourcePage.FieldExistsAndUsed(Constants.FileRefField) || !publishingPageTransformationInformation.SourcePage.FieldExistsAndUsed(Constants.FileLeafRefField))
            {
                LogError(LogStrings.Error_PageNotValidMissingFileRef, LogStrings.Heading_InputValidation);
                throw new ArgumentException(LogStrings.Error_PageNotValidMissingFileRef);
            }

            string pageType = publishingPageTransformationInformation.SourcePage.PageType();

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

            if (pageType.Equals("WikiPage", StringComparison.InvariantCultureIgnoreCase) || pageType.Equals("WebPartPage", StringComparison.InvariantCultureIgnoreCase))
            {
                LogError(LogStrings.Error_PageIsNotAPublishingPage, LogStrings.Heading_InputValidation);
                throw new ArgumentException(LogStrings.Error_PageIsNotAPublishingPage);
            }

            // Disable cross-farm item level permissions from copying
            CrossFarmTransformationValidation(publishingPageTransformationInformation);

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

                LogDebug(LogStrings.LoadingTargetClientContext, LogStrings.Heading_SharePointConnection);
                LoadClientObject(targetClientContext, true);

                SetAADTenantId(sourceClientContext, targetClientContext);

                if (sourceClientContext.Site.Id.Equals(targetClientContext.Site.Id))
                {
                    // Oops, seems source and target point to the same site collection...that's a no go for publishing portal page transformation!                
                    LogError(LogStrings.Error_SameSiteTransferNoAllowedForPublishingPages, LogStrings.Heading_SharePointConnection);
                    throw new ArgumentNullException(LogStrings.Error_SameSiteTransferNoAllowedForPublishingPages);
                }

                LogInfo($"{targetClientContext.Web.GetUrl()}", LogStrings.Heading_Summary, LogEntrySignificance.TargetSiteUrl);

                // Need to add further validation for target template
                if (targetClientContext.Web.WebTemplate != "SITEPAGEPUBLISHING" && targetClientContext.Web.WebTemplate != "STS" && targetClientContext.Web.WebTemplate != "GROUP")
                {

                    LogError(LogStrings.Error_CrossSiteTransferTargetsNonModernSite);
                    throw new ArgumentException(LogStrings.Error_CrossSiteTransferTargetsNonModernSite, LogStrings.Heading_SharePointConnection);
                }

                LogInfo($"{publishingPageTransformationInformation.SourcePage[Constants.FileRefField].ToString().ToLower()}", LogStrings.Heading_Summary, LogEntrySignificance.SourcePage);

#if DEBUG && MEASURE
            Stop("Telemetry");
#endif
                #endregion

                #region Page creation
                // Detect if the page is living inside a folder
                LogDebug(LogStrings.DetectIfPageIsInFolder, LogStrings.Heading_PageCreation);
                string pageFolder = "";

                // Get the publishing pages library name
                this.publishingPagesLibraryName = CacheManager.Instance.GetPublishingPagesLibraryName(this.sourceClientContext);

                if (publishingPageTransformationInformation.SourcePage.FieldExistsAndUsed(Constants.FileDirRefField))
                {
                    var fileRefFieldValue = publishingPageTransformationInformation.SourcePage[Constants.FileDirRefField].ToString().ToLower();

                    if (fileRefFieldValue.Contains($"/{this.publishingPagesLibraryName}"))
                    {
                        string pagesLibraryRelativeUrl = $"{sourceClientContext.Web.ServerRelativeUrl.TrimEnd(new[] { '/' })}/{this.publishingPagesLibraryName}";
                        pageFolder = fileRefFieldValue.Replace(pagesLibraryRelativeUrl.ToLower(), "").Trim();
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
                }
                publishingPageTransformationInformation.Folder = pageFolder;

                // If no targetname specified then we'll come up with one
                if (string.IsNullOrEmpty(publishingPageTransformationInformation.TargetPageName))
                {
                    LogInfo(LogStrings.CrossSiteInUseUsingOriginalFileName, LogStrings.Heading_PageCreation);
                    publishingPageTransformationInformation.TargetPageName = $"{publishingPageTransformationInformation.SourcePage[Constants.FileLeafRefField].ToString()}";
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
                var context = targetClientContext;

                try
                {
                    LogDebug(LogStrings.LoadingExistingPageIfExists, LogStrings.Heading_PageCreation);

                    // Just try to load the page in the fastest possible manner, we only want to see if the page exists or not
                    existingFile = Load(sourceClientContext, targetClientContext, publishingPageTransformationInformation, out pagesLibrary);
                    pageExists = true;
                }
                catch (Exception ex)
                {
                    if(ex is ArgumentException)
                    {
                        //Non-critical error generated 
                        LogInfo(LogStrings.CheckPageExistsError, LogStrings.Heading_PageCreation);
                    }
                    else
                    {
                        //Something else occurred
                        LogError(LogStrings.CheckPageExistsError, LogStrings.Heading_PageCreation, ex);
                    }
                }
                
#if DEBUG && MEASURE
            Stop("Load Page");
#endif

                if (pageExists)
                {
                    LogInfo(LogStrings.PageAlreadyExistsInTargetLocation, LogStrings.Heading_PageCreation);

                    if (!publishingPageTransformationInformation.Overwrite)
                    {
                        var message = $"{LogStrings.PageNotOverwriteIfExists}  {publishingPageTransformationInformation.TargetPageName}.";
                        LogError(message, LogStrings.Heading_PageCreation);
                        throw new ArgumentException(message);
                    }
                }

                // Create the client side page
                targetPage = context.Web.AddClientSidePage($"{publishingPageTransformationInformation.Folder}{publishingPageTransformationInformation.TargetPageName}");
                LogInfo($"{LogStrings.ModernPageCreated} ", LogStrings.Heading_PageCreation);
                #endregion

                LogInfo(LogStrings.TransformSourcePageAsArticlePage, LogStrings.Heading_ArticlePageHandling);

                #region Analysis of the source page
#if DEBUG && MEASURE
                Start();
#endif
                // Analyze the source page
                Tuple<Pages.PageLayout, List<WebPartEntity>> pageData = null;

                LogInfo($"{LogStrings.TransformSourcePageIsPublishingPage} - {LogStrings.TransformSourcePageAnalysing}", LogStrings.Heading_ArticlePageHandling);

                // Grab the pagelayout mapping to use:
                var pageLayoutMappingModel = new PageLayoutManager(this.RegisteredLogObservers).GetPageLayoutMappingModel(this.publishingPageTransformation, publishingPageTransformationInformation.SourcePage);

                var spVersion = GetVersion(sourceClientContext);

                if (spVersion == SPVersion.SP2010 || spVersion == SPVersion.SP2013Legacy || spVersion == SPVersion.SP2016Legacy)
                {
                    pageData = new PublishingPageOnPremises(publishingPageTransformationInformation.SourcePage, pageTransformation, this.publishingPageTransformation, publishingPageTransformationInformation as BaseTransformationInformation, targetContext: targetClientContext, logObservers: base.RegisteredLogObservers).Analyze(pageLayoutMappingModel);
                }
                else {
                    pageData = new PublishingPage(publishingPageTransformationInformation.SourcePage, pageTransformation, this.publishingPageTransformation, publishingPageTransformationInformation as BaseTransformationInformation, targetContext: targetClientContext, logObservers: base.RegisteredLogObservers).Analyze(pageLayoutMappingModel);
                }


                // Wiki content can contain embedded images and videos, which is not supported by the target RTE...split wiki text blocks so the transformator can handle the images and videos as separate web parts
                LogInfo(LogStrings.WikiTextContainsImagesVideosReferences, LogStrings.Heading_ArticlePageHandling);
                pageData = new Tuple<Pages.PageLayout, List<WebPartEntity>>(pageData.Item1, new WikiHtmlTransformator(this.sourceClientContext, targetPage, publishingPageTransformationInformation as BaseTransformationInformation, base.RegisteredLogObservers).TransformPlusSplit(pageData.Item2, publishingPageTransformationInformation.HandleWikiImagesAndVideos, publishingPageTransformationInformation.AddTableListImageAsImageWebPart));

#if DEBUG && MEASURE
                Stop("Analyze page");
#endif
                #endregion

                #region Page title configuration
#if DEBUG && MEASURE
                Start();
#endif
                // Set page title
                SetPageTitle(publishingPageTransformationInformation, targetPage);

                if (publishingPageTransformationInformation.PageTitleOverride != null)
                {
                    var title = publishingPageTransformationInformation.PageTitleOverride(targetPage.PageTitle);
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
                bool useCustomLayoutTransformator = false;
                if (publishingPageTransformationInformation.LayoutTransformatorOverride != null)
                {
                    LogInfo(LogStrings.TransformLayoutTransformatorOverride, LogStrings.Heading_ArticlePageHandling);
                    layoutTransformator = publishingPageTransformationInformation.LayoutTransformatorOverride(targetPage);
                    useCustomLayoutTransformator = true;
                }

                // Apply the layout to the page
                layoutTransformator.Transform(pageData);

                // If needed call the specific publishing page layout transformator
                if (pageData.Item1 == Pages.PageLayout.PublishingPage_AutoDetect && !useCustomLayoutTransformator)
                {
                    // Call out the specific publishing layout transformator implementation
                    PublishingLayoutTransformator publishingLayoutTransformator = new PublishingLayoutTransformator(targetPage, base.RegisteredLogObservers);
                    publishingLayoutTransformator.Transform(pageData);
                }

#if DEBUG && MEASURE
                Stop("Page layout");
#endif
                #endregion

                #region Content transformation

                LogDebug(LogStrings.PreparingContentTransformation, LogStrings.Heading_ArticlePageHandling);

#if DEBUG && MEASURE
                Start();
#endif
                // Use the default content transformator
                IContentTransformator contentTransformator = new ContentTransformator(sourceClientContext, targetPage, pageTransformation, publishingPageTransformationInformation as BaseTransformationInformation, base.RegisteredLogObservers);

                // Do we have an override?
                if (publishingPageTransformationInformation.ContentTransformatorOverride != null)
                {
                    LogInfo(LogStrings.TransformUsingContentTransformerOverride, LogStrings.Heading_ArticlePageHandling);

                    contentTransformator = publishingPageTransformationInformation.ContentTransformatorOverride(targetPage, pageTransformation);
                }

                LogInfo(LogStrings.TransformingContentStart, LogStrings.Heading_ArticlePageHandling);

                // Run the content transformator
                contentTransformator.Transform(pageData.Item2.Where(c => !c.IsClosed).ToList());

                LogInfo(LogStrings.TransformingContentEnd, LogStrings.Heading_ArticlePageHandling);
#if DEBUG && MEASURE
                Stop("Content transformation");
#endif
                #endregion

                #region Configure header for target page
#if DEBUG && MEASURE
                Start();
#endif
                PublishingPageHeaderTransformator headerTransformator = new PublishingPageHeaderTransformator(publishingPageTransformationInformation, sourceClientContext, targetClientContext, this.publishingPageTransformation, base.RegisteredLogObservers);
                headerTransformator.TransformHeader(ref targetPage);

#if DEBUG && MEASURE
                Stop("Target page header");
#endif
                #endregion

                #region Text/Section/Column cleanup
                // Drop "empty" text parts. Wiki pages tend to have a lot of text parts just containing div's and BR's...no point in keep those as they generate to much whitespace
                RemoveEmptyTextParts(targetPage);

                // Remove empty sections and columns to optimize screen real estate
                if (publishingPageTransformationInformation.RemoveEmptySectionsAndColumns)
                {
                    RemoveEmptySectionsAndColumns(targetPage);
                }
                #endregion

                #region Page persisting + permissions

                #region Save the page
#if DEBUG && MEASURE
            Start();
#endif
                // Persist the client side page
                var pageName = $"{publishingPageTransformationInformation.Folder}{publishingPageTransformationInformation.TargetPageName}";
                targetPage.Save(pageName);
                LogInfo($"{LogStrings.TransformSavedPageInCrossSiteCollection}: {pageName}", LogStrings.Heading_ArticlePageHandling);

#if DEBUG && MEASURE
            Stop("Persist page");
#endif
                #endregion

                #region Page metadata handling
                PublishingMetadataTransformator publishingMetadataTransformator = new PublishingMetadataTransformator(publishingPageTransformationInformation, sourceClientContext, targetClientContext, targetPage, pageLayoutMappingModel, this.publishingPageTransformation, base.RegisteredLogObservers);
                publishingMetadataTransformator.Transform();
                #endregion

                #region Permission handling
                ListItemPermission listItemPermissionsToKeep = null;
                if (publishingPageTransformationInformation.KeepPageSpecificPermissions)
                {
#if DEBUG && MEASURE
                Start();
#endif
                    // Check if we do have item level permissions we want to take over
                    listItemPermissionsToKeep = GetItemLevelPermissions(true, pagesLibrary, publishingPageTransformationInformation.SourcePage, targetPage.PageListItem);

                    // When creating the page in another site collection we'll always want to copy item level permissions if specified
                    ApplyItemLevelPermissions(true, targetPage.PageListItem, listItemPermissionsToKeep);
#if DEBUG && MEASURE
                Stop("Permission handling");
#endif
                }
                #endregion

                #region Page Publishing
                // Tag the file with a page modernization version stamp
                string serverRelativePathForModernPage = ReturnModernPageServerRelativeUrl(publishingPageTransformationInformation);
                bool pageListItemWasReloaded = false;
                try
                {
                    var targetPageFile = context.Web.GetFileByServerRelativeUrl(serverRelativePathForModernPage);
                    context.Load(targetPageFile, p => p.Properties);
                    targetPageFile.Properties["sharepointpnp_pagemodernization"] = this.version;
                    targetPageFile.Update();

                    if (publishingPageTransformationInformation.PublishCreatedPage)
                    {
                        // Try to publish, if publish is not needed/possible (e.g. when no minor/major versioning set) then this will return an error that we'll be ignoring
                        targetPageFile.Publish("Page modernization initial publish");
                    }

                    // Ensure we've the most recent page list item loaded, must be last statement before calling ExecuteQuery
                    context.Load(targetPage.PageListItem);
                    // Send both the property update and publish as a single operation to SharePoint
                    context.ExecuteQueryRetry();
                    pageListItemWasReloaded = true;
                }
                catch (Exception ex)
                {
                    // Eat exceptions as this is not critical for the generated page
                    LogWarning(LogStrings.Warning_NonCriticalErrorDuringVersionStampAndPublish, LogStrings.Heading_ArticlePageHandling);                    
                }

                // Update flags field to indicate this is a "migrated" page
                try
                {
                    // If for some reason the reload batched with the previous request did not finish then do it again
                    if (!pageListItemWasReloaded)
                    {
                        context.Load(targetPage.PageListItem);
                        context.ExecuteQueryRetry();
                    }

                    // Only perform the update when the field was not yet set
                    bool skipSettingMigratedFromServerRendered = false;
                    if (targetPage.PageListItem[Constants.SPSitePageFlagsField] != null)
                    {
                        skipSettingMigratedFromServerRendered = (targetPage.PageListItem[Constants.SPSitePageFlagsField] as string[]).Contains("MigratedFromServerRendered");
                    }

                    if (!skipSettingMigratedFromServerRendered)
                    {
                        targetPage.PageListItem[Constants.SPSitePageFlagsField] = ";#MigratedFromServerRendered;#";
                        targetPage.PageListItem.Update();
                        context.Load(targetPage.PageListItem);
                        context.ExecuteQueryRetry();
                    }
                }
                catch (Exception ex)
                {
                    // Eat any exception
                }

                // Disable page comments on the create page, if needed
                if (publishingPageTransformationInformation.DisablePageComments)
                {
                    targetPage.DisableComments();
                    LogInfo(LogStrings.TransformDisablePageComments, LogStrings.Heading_ArticlePageHandling);
                }
                #endregion

                #region Telemetry
                if (!publishingPageTransformationInformation.SkipTelemetry && this.pageTelemetry != null)
                {
                    TimeSpan duration = DateTime.Now.Subtract(transformationStartDateTime);
                    this.pageTelemetry.LogTransformationDone(duration, pageType, publishingPageTransformationInformation);
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

        #region Helper methods

        private string ReturnModernPageServerRelativeUrl(PublishingPageTransformationInformation publishingPageTransformationInformation)
        {
            string originalSourcePageName = publishingPageTransformationInformation.TargetPageName.ToLower();
            string sourcePath = publishingPageTransformationInformation.SourcePage[Constants.FileRefField].ToString().ToLower().Replace(publishingPageTransformationInformation.SourcePage[Constants.FileLeafRefField].ToString().ToLower(), "");
            string targetPath;

            // Cross site collection transfer, new page always takes the name of the old page
            if (!sourcePath.Contains($"/{this.publishingPagesLibraryName}"))
            {
                // Source file was living outside of the site pages library
                targetPath = sourcePath.Replace(sourceClientContext.Web.ServerRelativeUrl.ToLower(), "");
                targetPath = $"{targetClientContext.Web.ServerRelativeUrl.ToLower()}/sitepages{targetPath}";
            }
            else
            {
                // Page was living inside the sitepages library
                targetPath = sourcePath.Replace($"{sourceClientContext.Web.ServerRelativeUrl}/{this.publishingPagesLibraryName}".ToLower(), "");
                targetPath = $"{targetClientContext.Web.ServerRelativeUrl.ToLower()}/sitepages{targetPath}";
            }

            string returnUrl = $"{targetPath}{originalSourcePageName}";
          
            LogInfo($"{returnUrl}", LogStrings.Heading_Summary, LogEntrySignificance.TargetPage);
            return returnUrl;
        }

        private void SetPageTitle(PublishingPageTransformationInformation publishingPageTransformationInformation, ClientSidePage targetPage)
        {
            string titleValue = "";
            if (publishingPageTransformationInformation.SourcePage.FieldExistsAndUsed(Constants.TitleField))
            {
                titleValue = publishingPageTransformationInformation.SourcePage[Constants.TitleField].ToString();
            }
            else if (publishingPageTransformationInformation.SourcePage.FieldExistsAndUsed(Constants.FileLeafRefField))
            {
                titleValue = Path.GetFileNameWithoutExtension((publishingPageTransformationInformation.SourcePage[Constants.FileLeafRefField].ToString()));
            }

            if (!string.IsNullOrEmpty(titleValue))
            {
                titleValue = titleValue.First().ToString().ToUpper() + titleValue.Substring(1);
                targetPage.PageTitle = titleValue;
                LogInfo($"{LogStrings.TransformPageModernTitle} {titleValue}", LogStrings.Heading_SetPageTitle);
            }
        }

        private Microsoft.SharePoint.Client.File Load(ClientContext sourceContext, ClientContext targetContext, PublishingPageTransformationInformation publishingPageTransformationInformation, out List pagesLibrary)
        {
            sourceContext.Web.EnsureProperty(w => w.ServerRelativeUrl);
            sourceContext.Site.EnsureProperty(w => w.Url); 
            targetContext.Web.EnsureProperty(w => w.ServerRelativeUrl);

            // Load the pages library and page file (if exists) in one go 
            pagesLibrary = sourceContext.Web.GetListById(Guid.Parse(sourceContext.Web.GetPagesLibraryId())); 

            sourceContext.Web.Context.Load(pagesLibrary, l => l.DefaultViewUrl, l => l.Id, l => l.BaseTemplate, l => l.OnQuickLaunch, l => l.DefaultViewUrl, l => l.Title,
                                              l => l.Hidden, l => l.EffectiveBasePermissions, l => l.RootFolder, l => l.RootFolder.ServerRelativeUrl);

            if (publishingPageTransformationInformation.KeepPageSpecificPermissions)
            {
                sourceContext.Load(publishingPageTransformationInformation.SourcePage, p => p.HasUniqueRoleAssignments);
            }
            try
            {
                sourceClientContext.ExecuteQueryRetry();
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

            var sitePagesServerRelativeUrl = UrlUtility.Combine(targetClientContext.Web.ServerRelativeUrl, "sitepages");
            var file = targetClientContext.Web.GetFileByServerRelativeUrl($"{sitePagesServerRelativeUrl}/{publishingPageTransformationInformation.Folder}{publishingPageTransformationInformation.TargetPageName}");
            targetClientContext.Web.Context.Load(file, f => f.Exists, f => f.ListItemAllFields);
            targetClientContext.ExecuteQueryRetry();

            if (pagesLibrary == null)
            {
                LogError(LogStrings.Error_MissingSitePagesLibrary, LogStrings.Heading_Load);
                throw new ArgumentNullException(LogStrings.Error_MissingSitePagesLibrary);
            }

            if (!file.Exists)
            {
                LogInfo(LogStrings.TransformPageDoesNotExistInWeb, LogStrings.Heading_Load); //Not an error this is a check
                throw new ArgumentException($"{publishingPageTransformationInformation.TargetPageName} - {LogStrings.TransformPageDoesNotExistInWeb}");
            }

            return file;
        }

        /// <summary>
        /// Use reflection to read the object properties and detail the values
        /// </summary>
        /// <param name="pti">PageTransformationInformation object</param>
        /// <returns></returns>
        private List<LogEntry> DetailSettingsAsLogEntries(PublishingPageTransformationInformation pti)
        {
            List<LogEntry> logs = new List<LogEntry>();

            try
            {

                var properties = pti.GetType().GetProperties();
                foreach (var property in properties)
                {
                    if (property.PropertyType == typeof(String) ||
                        property.PropertyType == typeof(bool))
                    {
                        logs.Add(new LogEntry()
                        {
                            Heading = LogStrings.Heading_PageTransformationInfomation,
                            Message = $"{property.Name.FormatAsFriendlyTitle()} {LogStrings.KeyValueSeperatorToken} {property.GetValue(pti)}"
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                logs.Add(new LogEntry() { Message = "Failed to convert object properties for reporting", Exception = ex, Heading = LogStrings.Heading_PageTransformationInfomation });
            }

            return logs;

        }
        #endregion

    }
}
