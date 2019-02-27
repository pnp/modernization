using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using OfficeDevPnP.Core.Pages;
using OfficeDevPnP.Core.Utilities;
using SharePointPnP.Modernization.Framework.Cache;
using SharePointPnP.Modernization.Framework.Entities;
using SharePointPnP.Modernization.Framework.Pages;
using SharePointPnP.Modernization.Framework.Telemetry;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml.Serialization;

namespace SharePointPnP.Modernization.Framework.Transform
{

    /// <summary>
    /// Transforms a classic wiki/webpart page into a modern client side page
    /// </summary>
    public class PageTransformator
    {
        private ClientContext sourceClientContext;
        private ClientContext targetClientContext;
        private PageTransformation pageTransformation;
        private string version = "undefined";
        private PageTelemetry pageTelemetry;
        private Stopwatch watch;
        private const string ExecutionLog = "execution.csv";

        #region Construction

        /// <summary>
        /// Creates a page transformator instance with a target destination of a target web e.g. Modern/Communication Site
        /// </summary>
        /// <param name="clientContext">ClientContext of the site holding the page</param>
        /// <param name="targetClientContext">ClientContext of the site that will receive the modernized page</param>
        public PageTransformator(ClientContext sourceClientContext, ClientContext targetClientContext) : this(sourceClientContext, targetClientContext, "webpartmapping.xml")
        {

        }

        /// <summary>
        /// Creates a page transformator instance
        /// </summary>
        /// <param name="clientContext">ClientContext of the site holding the page</param>
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

            this.version = PageTransformator.GetVersion();
            this.pageTelemetry = new PageTelemetry(version);

            // Load xml mapping data
            XmlSerializer xmlMapping = new XmlSerializer(typeof(PageTransformation));
            using (var stream = new FileStream(pageTransformationFile, FileMode.Open))
            {
                this.pageTransformation = (PageTransformation)xmlMapping.Deserialize(stream);
            }
        }

        /// <summary>
        /// Creates a page transformator instance
        /// </summary>
        /// <param name="clientContext">ClientContext of the site holding the page</param>
        /// <param name="pageTransformationModel">Page transformation model</param>
        public PageTransformator(ClientContext sourceClientContext, PageTransformation pageTransformationModel)
        {

#if DEBUG && MEASURE
            InitMeasurement();
#endif

            this.sourceClientContext = sourceClientContext;
            this.version = PageTransformator.GetVersion();
            this.pageTelemetry = new PageTelemetry(version);

            this.pageTransformation = pageTransformationModel;
        }
        #endregion

        /// <summary>
        /// Transform the page
        /// </summary>
        /// <param name="pageTransformationInformation">Information about the page to transform</param>
        /// <returns>The path to created modern page</returns>
        public string Transform(PageTransformationInformation pageTransformationInformation)
        {
            #region Check for Target Site Context
            var hasTargetContext = targetClientContext != null;
            #endregion

            #region Input validation
            if (pageTransformationInformation.SourcePage == null)
            {
                throw new ArgumentNullException("SourcePage cannot be null");
            }

            // Validate page and it's eligibility for transformation
            if (!pageTransformationInformation.SourcePage.FieldExistsAndUsed(Constants.FileRefField) || !pageTransformationInformation.SourcePage.FieldExistsAndUsed(Constants.FileLeafRefField))
            {
                throw new ArgumentException("Page is not valid due to missing FileRef or FileLeafRef value");
            }

            string pageType = pageTransformationInformation.SourcePage.PageType();

            if (pageType.Equals("ClientSidePage", StringComparison.InvariantCultureIgnoreCase))
            {
                throw new ArgumentException("Page already is a modern client side page...no need to transform it.");
            }

            if (pageType.Equals("AspxPage", StringComparison.InvariantCultureIgnoreCase))
            {
                throw new ArgumentException("Page is an basic aspx page...can't currently transform that one, sorry!");
            }

            if (pageType.Equals("PublishingPage", StringComparison.InvariantCultureIgnoreCase))
            {
                throw new ArgumentException("Page transformation for publishing pages is currently not supported.");
            }

            if (hasTargetContext)
            {
                // If we're transforming into another site collection the "revert to old page" model does not exist as the 
                // old page is not present in there. Also adding the page transformation banner does not make sense for the same reason
                if (pageTransformationInformation.ModernizationCenterInformation != null && pageTransformationInformation.ModernizationCenterInformation.AddPageAcceptBanner)
                {
                    throw new ArgumentException("Page transformation towards a different site collection cannot use the page accept banner.");
                }
            }

            #endregion

            #region Telemetry
#if DEBUG && MEASURE
            Start();
#endif            
            DateTime transformationStartDateTime = DateTime.Now;

            LoadClientObject(sourceClientContext);
            if (hasTargetContext)
            {
                LoadClientObject(targetClientContext);
            }

            if (hasTargetContext)
            {
                if (sourceClientContext.Site.Id.Equals(targetClientContext.Site.Id))
                {
                    // Oops, seems source and target point to the same site collection...switch back the "source only" mode
                    targetClientContext = null;
                    hasTargetContext = false;
                }
                else
                {
                    // Ensure that the newly created page in the other site collection gets the same name as the source page
                    pageTransformationInformation.TargetPageTakesSourcePageName = true;
                }
            }

            // Need to add further validation for target template
            if (hasTargetContext &&
               (targetClientContext.Web.WebTemplate != "SITEPAGEPUBLISHING" && targetClientContext.Web.WebTemplate != "STS" && targetClientContext.Web.WebTemplate != "GROUP"))
            {
                throw new ArgumentException("Page transformation for targeting non-modern sites is currently not supported.");
            }

#if DEBUG && MEASURE
            Stop("Telemetry");
#endif
            #endregion

            #region Page creation
            // Detect if the page is living inside a folder
            string pageFolder = "";
            if (pageTransformationInformation.SourcePage.FieldExistsAndUsed(Constants.FileDirRefField))
            {
                var fileRefFieldValue = pageTransformationInformation.SourcePage[Constants.FileDirRefField].ToString();
                pageFolder = fileRefFieldValue.Replace($"{sourceClientContext.Web.ServerRelativeUrl}/SitePages", "").Trim();

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
                }
            }
            pageTransformationInformation.Folder = pageFolder;

            // If no targetname specified then we'll come up with one
            if (string.IsNullOrEmpty(pageTransformationInformation.TargetPageName))
            {
                if (string.IsNullOrEmpty(pageTransformationInformation.TargetPagePrefix))
                {
                    pageTransformationInformation.SetDefaultTargetPagePrefix();
                }

                if (hasTargetContext)
                {
                    pageTransformationInformation.TargetPageName = $"{pageTransformationInformation.SourcePage[Constants.FileLeafRefField].ToString()}";
                }
                else
                {
                    pageTransformationInformation.TargetPageName = $"{pageTransformationInformation.TargetPagePrefix}{pageTransformationInformation.SourcePage[Constants.FileLeafRefField].ToString()}";
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
                // Just try to load the page in the fastest possible manner, we only want to see if the page exists or not
                existingFile = Load(sourceClientContext, pageTransformationInformation, out pagesLibrary);
                pageExists = true;
            }
            catch (ArgumentException) { }
#if DEBUG && MEASURE
            Stop("Load Page");
#endif            

            if (pageExists)
            {
                if (!pageTransformationInformation.Overwrite)
                {
                    throw new ArgumentException($"There already exists a page with name {pageTransformationInformation.TargetPageName}.");
                }
            }

            // Create the client side page

            targetPage = context.Web.AddClientSidePage($"{pageTransformationInformation.Folder}{pageTransformationInformation.TargetPageName}");
            #endregion

            #region Home page handling
#if DEBUG && MEASURE
            Start();
#endif
            bool replacedByOOBHomePage = false;
            // Check if the transformed page is the web's home page
            if (sourceClientContext.Web.RootFolder.IsPropertyAvailable("WelcomePage") && !string.IsNullOrEmpty(sourceClientContext.Web.RootFolder.WelcomePage))
            {
                var homePageUrl = sourceClientContext.Web.RootFolder.WelcomePage;
                var homepageName = Path.GetFileName(sourceClientContext.Web.RootFolder.WelcomePage);
                if (homepageName.Equals(pageTransformationInformation.SourcePage[Constants.FileLeafRefField].ToString(), StringComparison.InvariantCultureIgnoreCase))
                {
                    targetPage.LayoutType = ClientSidePageLayoutType.Home;
                    if (pageTransformationInformation.ReplaceHomePageWithDefaultHomePage)
                    {
                        targetPage.KeepDefaultWebParts = true;
                        replacedByOOBHomePage = true;
                    }
                }
            }
#if DEBUG && MEASURE
            Stop("Home page handling");
#endif            
            #endregion

            #region Article page handling
            if (!replacedByOOBHomePage)
            {
                #region Configure header from target page
#if DEBUG && MEASURE
                Start();
#endif            
                if (pageTransformationInformation.PageHeader == null || pageTransformationInformation.PageHeader.Type == ClientSidePageHeaderType.None)
                {
                    targetPage.RemovePageHeader();
                }
                else if (pageTransformationInformation.PageHeader.Type == ClientSidePageHeaderType.Default)
                {
                    targetPage.SetDefaultPageHeader();
                }
                else if (pageTransformationInformation.PageHeader.Type == ClientSidePageHeaderType.Custom)
                {
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
                    pageData = new WikiPage(pageTransformationInformation.SourcePage, pageTransformation).Analyze();

                    // Wiki pages can contain embedded images and videos, which is not supported by the target RTE...split wiki text blocks so the transformator can handle the images and videos as separate web parts
                    pageData = new Tuple<PageLayout, List<WebPartEntity>>(pageData.Item1, new WikiTransformatorSimple().TransformPlusSplit(pageData.Item2, pageTransformationInformation.HandleWikiImagesAndVideos));
                }
                else if (pageType.Equals("WebPartPage", StringComparison.InvariantCultureIgnoreCase))
                {
                    pageData = new WebPartPage(pageTransformationInformation.SourcePage, pageTransformation).Analyze(true);
                }
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
                            targetPage.PageTitle = titleBarWebPart.Properties["HeaderTitle"];
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
                    targetPage.PageTitle = pageTransformationInformation.PageTitleOverride(targetPage.PageTitle);
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
                    layoutTransformator = pageTransformationInformation.LayoutTransformatorOverride(targetPage);
                }

                // Apply the layout to the page
                layoutTransformator.Transform(pageData.Item1);
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
                        foreach(var webpart in pageData.Item2)
                        {
                            webpart.Row = webpart.Row + 1;
                        }


                        var sourcePageUrl = pageTransformationInformation.SourcePage[Constants.FileRefField].ToString();
                        var orginalSourcePageName = pageTransformationInformation.SourcePage[Constants.FileLeafRefField].ToString();
                        Uri host = new Uri(sourceClientContext.Web.Url);

                        string path = $"{host.Scheme}://{host.DnsSafeHost}{sourcePageUrl.Replace(pageTransformationInformation.SourcePage[Constants.FileLeafRefField].ToString(), "")}";

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
#if DEBUG && MEASURE
                        Stop("Page Banner");
#endif
                    }
                }
                #endregion  

                #region Content transformation
#if DEBUG && MEASURE
                Start();
#endif            
                // Use the default content transformator
                IContentTransformator contentTransformator = new ContentTransformator(targetPage, pageTransformation);

                // Do we have an override?
                if (pageTransformationInformation.ContentTransformatorOverride != null)
                {
                    contentTransformator = pageTransformationInformation.ContentTransformatorOverride(targetPage, pageTransformation);
                }

                // Run the content transformator
                contentTransformator.Transform(pageData.Item2);
#if DEBUG && MEASURE
                Stop("Content transformation");
#endif
                #endregion

                #region Section/column cleanup
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
                targetPage.Save($"{pageTransformationInformation.Folder}{pageTransformationInformation.TargetPageName}");
            }
            else
            { 
                targetPage.Save($"{pageTransformationInformation.Folder}{pageTransformationInformation.TargetPageName}", existingFile, pagesLibrary);
            }

            // Tag the file with a page modernization version stamp
            try
            {
                string path = pageTransformationInformation.SourcePage[Constants.FileRefField].ToString().Replace(pageTransformationInformation.SourcePage[Constants.FileLeafRefField].ToString(), "");
                var targetPageUrl = $"{path}{pageTransformationInformation.TargetPageName}";
                var targetPageFile = this.sourceClientContext.Web.GetFileByServerRelativeUrl(targetPageUrl);
                this.sourceClientContext.Load(targetPageFile, p => p.Properties);
                targetPageFile.Properties["sharepointpnp_pagemodernization"] = this.version;
                targetPageFile.Update();

                // Try to publish, if publish is not needed then this will return an error that we'll be ignoring
                targetPageFile.Publish("Page modernization initial publish");

                this.sourceClientContext.ExecuteQueryRetry();
                if (hasTargetContext)
                {
                    this.targetClientContext.ExecuteQueryRetry();
                }
            }
            catch (Exception ex)
            {
                // Eat exceptions as this is not critical for the generated page
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
                this.pageTelemetry.LogTransformationDone(duration);
                this.pageTelemetry.Flush();
            }
            #endregion

            #region Return final page url
            if(!pageTransformationInformation.TargetPageTakesSourcePageName)
            {
                string path = pageTransformationInformation.SourcePage[Constants.FileRefField].ToString().Replace(pageTransformationInformation.SourcePage[Constants.FileLeafRefField].ToString(), "");
                var targetPageUrl = $"{path}{pageTransformationInformation.TargetPageName}";
                return targetPageUrl;
            }
            else
            {
                return pageTransformationInformation.SourcePage[Constants.FileRefField].ToString();
            }
            #endregion

            #endregion
        }

        /// <summary>
        /// Performs the logic needed to swap a genered Migrated_Page.aspx to Page.aspx and then Page.aspx to Old_Page.aspx
        /// </summary>
        /// <param name="pageTransformationInformation">Information about the page to transform</param>
        public void SwapPages(PageTransformationInformation pageTransformationInformation, ListItemPermission listItemPermissionsToKeep)
        {
            var sourcePageUrl = pageTransformationInformation.SourcePage[Constants.FileRefField].ToString();
            var orginalSourcePageName = pageTransformationInformation.SourcePage[Constants.FileLeafRefField].ToString();

            string path = sourcePageUrl.Replace(pageTransformationInformation.SourcePage[Constants.FileLeafRefField].ToString(), "");

            var sourcePage = this.sourceClientContext.Web.GetFileByServerRelativeUrl(sourcePageUrl);
            this.sourceClientContext.Load(sourcePage);
            this.sourceClientContext.ExecuteQueryRetry();

            if (string.IsNullOrEmpty(pageTransformationInformation.SourcePagePrefix))
            {
                pageTransformationInformation.SetDefaultSourcePagePrefix();
            }
            var newSourcePageUrl = $"{pageTransformationInformation.SourcePagePrefix}{pageTransformationInformation.SourcePage[Constants.FileLeafRefField].ToString()}";

            // Rename source page using the sourcepageprefix
            // STEP1: First copy the source page to a new name. We on purpose use CopyTo as we want to avoid that "linked" url's get 
            //        patched up during a MoveTo operation as that would also patch the url's in our new modern page
            sourcePage.CopyTo($"{path}{newSourcePageUrl}", true);
            this.sourceClientContext.ExecuteQueryRetry();

            // Restore the item level permissions on the copied page (if any)
            if (pageTransformationInformation.KeepPageSpecificPermissions && listItemPermissionsToKeep != null)
            {
                // load the copied target file
                var newSource = this.sourceClientContext.Web.GetFileByServerRelativeUrl($"{path}{newSourcePageUrl}");
                this.sourceClientContext.Load(newSource);
                this.sourceClientContext.Load(newSource.ListItemAllFields, p => p.RoleAssignments);
                this.sourceClientContext.ExecuteQueryRetry();

                // Reload source page
                ApplyItemLevelPermissions(false, newSource.ListItemAllFields, listItemPermissionsToKeep, alwaysBreakItemLevelPermissions: true);
            }

            //Load the created target page
            var targetPageUrl = $"{path}{pageTransformationInformation.TargetPageName}";
            var targetPageFile = this.sourceClientContext.Web.GetFileByServerRelativeUrl(targetPageUrl);
            this.sourceClientContext.Load(targetPageFile);
            this.sourceClientContext.ExecuteQueryRetry();

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
                    node.Url = $"{path}{newSourcePageUrl}";
                    node.Update();
                }
                foreach (var node in globalNavNodes)
                {
                    node.Url = $"{path}{newSourcePageUrl}";
                    node.Update();
                }
                this.sourceClientContext.ExecuteQueryRetry();
            }

            // STEP3: Now copy the created modern page over the original source page, at this point the new page has the same name as the original page had before transformation
            targetPageFile.CopyTo($"{path}{orginalSourcePageName}", true);
            this.sourceClientContext.ExecuteQueryRetry();

            // Apply the item level permissions on the final page (if any)
            if (pageTransformationInformation.KeepPageSpecificPermissions && listItemPermissionsToKeep != null)
            {
                // load the copied target file
                var newTarget = this.sourceClientContext.Web.GetFileByServerRelativeUrl($"{path}{orginalSourcePageName}");
                this.sourceClientContext.Load(newTarget);
                this.sourceClientContext.Load(newTarget.ListItemAllFields, p => p.RoleAssignments);
                this.sourceClientContext.ExecuteQueryRetry();

                ApplyItemLevelPermissions(false, newTarget.ListItemAllFields, listItemPermissionsToKeep, alwaysBreakItemLevelPermissions: true);
            }

            // STEP4: Finish with restoring the page navigation: update the navlinks to point back the original page name
            if (navWasFixed)
            {
                // Reload the navigation entries as did update them
                this.sourceClientContext.Web.Context.Load(this.sourceClientContext.Web, w => w.Navigation.QuickLaunch, w => w.Navigation.TopNavigationBar);
                this.sourceClientContext.Web.Context.ExecuteQueryRetry();

                currentNavigation = this.sourceClientContext.Web.Navigation.QuickLaunch;
                globalNavigation = this.sourceClientContext.Web.Navigation.TopNavigationBar;
                if (!string.IsNullOrEmpty($"{path}{newSourcePageUrl}"))
                {
                    currentNavNodes = currentNavigation.Where(n => n.Url.Equals($"{path}{newSourcePageUrl}", StringComparison.InvariantCultureIgnoreCase));
                    globalNavNodes = globalNavigation.Where(n => n.Url.Equals($"{path}{newSourcePageUrl}", StringComparison.InvariantCultureIgnoreCase));
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
            targetPageFile.DeleteObject();
            this.sourceClientContext.ExecuteQueryRetry();
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
        private void RemoveEmptySectionsAndColumns(ClientSidePage targetPage)
        {
            foreach (var section in targetPage.Sections.ToList())
            {
                // First remove all empty sections
                if (section.Controls.Count == 0)
                {
                    targetPage.Sections.Remove(section);
                }
            }

            // Remove empty columns
            foreach (var section in targetPage.Sections)
            {
                if (section.Type == CanvasSectionTemplate.TwoColumn ||
                    section.Type == CanvasSectionTemplate.TwoColumnLeft ||
                    section.Type == CanvasSectionTemplate.TwoColumnRight)
                {
                    var emptyColumn = section.Columns.Where(p => p.Controls.Count == 0).FirstOrDefault();
                    if (emptyColumn != null)
                    {
                        // drop the empty column and change to single column section
                        section.Columns.Remove(emptyColumn);
                        section.Type = CanvasSectionTemplate.OneColumn;
                        section.Columns.First().ResetColumn(0, 12);
                    }
                }
                else if (section.Type == CanvasSectionTemplate.ThreeColumn)
                {
                    var emptyColumns = section.Columns.Where(p => p.Controls.Count == 0);
                    if (emptyColumns != null)
                    {
                        if (emptyColumns.Any() && emptyColumns.Count() == 2)
                        {
                            // drop the two empty columns and change to single column section
                            foreach (var emptyColumn in emptyColumns.ToList())
                            {
                                section.Columns.Remove(emptyColumn);
                            }
                            section.Type = CanvasSectionTemplate.OneColumn;
                            section.Columns.First().ResetColumn(0, 12);
                        }
                        else if (emptyColumns.Any() && emptyColumns.Count() == 1)
                        {
                            // Remove the empty column and change to two column section
                            section.Columns.Remove(emptyColumns.First());
                            section.Type = CanvasSectionTemplate.TwoColumn;
                            int i = 0;
                            foreach (var column in section.Columns)
                            {
                                column.ResetColumn(i, 6);
                                i++;
                            }
                        }
                    }
                }
            }
        }

        private void ApplyItemLevelPermissions(bool hasTargetContext, ListItem item, ListItemPermission lip, bool alwaysBreakItemLevelPermissions = false)
        {
            if (lip == null || item == null)
            {
                return;
            }

            // Break permission inheritance on the item if not done yet
            if (alwaysBreakItemLevelPermissions || !item.HasUniqueRoleAssignments)
            {
                item.BreakRoleInheritance(false, false);
                //this.sourceClientContext.ExecuteQueryRetry();
                item.Context.ExecuteQueryRetry();
            }

            if (hasTargetContext)
            {
                // Ensure principals are available in the target site
                Dictionary<string, Principal> targetPrincipals = new Dictionary<string, Principal>(lip.Principals.Count);
                foreach (var principal in lip.Principals)
                {
                    var targetPrincipal = GetPrincipal(this.targetClientContext.Web, principal.Key);
                    if (targetPrincipal != null)
                    {
                        if (!targetPrincipals.ContainsKey(principal.Key))
                        {
                            targetPrincipals.Add(principal.Key, targetPrincipal);
                        }
                    }
                }

                // Assign item level permissions          
                foreach (var roleAssignment in lip.RoleAssignments)
                {
                    if (targetPrincipals.TryGetValue(roleAssignment.Member.LoginName, out Principal principal))
                    {
                        var roleDefinitionBindingCollection = new RoleDefinitionBindingCollection(this.targetClientContext);
                        foreach (var roleDef in roleAssignment.RoleDefinitionBindings)
                        {
                            var targetRoleDef = this.targetClientContext.Web.RoleDefinitions.GetByName(roleDef.Name);
                            if (targetRoleDef != null)
                            {
                                roleDefinitionBindingCollection.Add(targetRoleDef);
                            }
                        }
                        item.RoleAssignments.Add(principal, roleDefinitionBindingCollection);
                    }
                }

                this.targetClientContext.ExecuteQueryRetry();
            }
            else
            {
                // Assign item level permissions
                foreach (var roleAssignment in lip.RoleAssignments)
                {
                    if (lip.Principals.TryGetValue(roleAssignment.Member.LoginName, out Principal principal))
                    {
                        var roleDefinitionBindingCollection = new RoleDefinitionBindingCollection(this.sourceClientContext);
                        foreach (var roleDef in roleAssignment.RoleDefinitionBindings)
                        {
                            roleDefinitionBindingCollection.Add(roleDef);
                        }

                        item.RoleAssignments.Add(principal, roleDefinitionBindingCollection);
                    }
                }

                this.sourceClientContext.ExecuteQueryRetry();
            }
        }


        private ListItemPermission GetItemLevelPermissions(bool hasTargetContext, List pagesLibrary, ListItem source, ListItem target)
        {
            ListItemPermission lip = null;

            if (source.HasUniqueRoleAssignments)
            {
                // You need to have the ManagePermissions permission before item level permissions can be copied
                if (pagesLibrary.EffectiveBasePermissions.Has(PermissionKind.ManagePermissions))
                {
                    // Copy the unique permissions from source to target
                    // Get the unique permissions
                    this.sourceClientContext.Load(source, a => a.EffectiveBasePermissions, a => a.RoleAssignments.Include(roleAsg => roleAsg.Member.LoginName,
                        roleAsg => roleAsg.RoleDefinitionBindings.Include(roleDef => roleDef.Name, roleDef => roleDef.Description)));
                    this.sourceClientContext.ExecuteQueryRetry();

                    if (source.EffectiveBasePermissions.Has(PermissionKind.ManagePermissions))
                    {
                        // Load the site groups
                        this.sourceClientContext.Load(this.sourceClientContext.Web.SiteGroups, p => p.Include(g => g.LoginName));

                        // Get target page information
                        if (hasTargetContext)
                        {
                            this.targetClientContext.Load(target, p => p.HasUniqueRoleAssignments, p => p.RoleAssignments);
                            this.targetClientContext.Load(this.targetClientContext.Web, p => p.RoleDefinitions);
                        }
                        else
                        {
                            this.sourceClientContext.Load(target, p => p.HasUniqueRoleAssignments, p => p.RoleAssignments);
                        }

                        this.sourceClientContext.ExecuteQueryRetry();

                        if (hasTargetContext)
                        {
                            this.targetClientContext.ExecuteQueryRetry();
                        }

                        Dictionary<string, Principal> principals = new Dictionary<string, Principal>(10);
                        lip = new ListItemPermission()
                        {
                            RoleAssignments = source.RoleAssignments,
                            Principals = principals
                        };

                        // Apply new permissions
                        foreach (var roleAssignment in source.RoleAssignments)
                        {
                            var principal = GetPrincipal(this.sourceClientContext.Web, roleAssignment.Member.LoginName);
                            if (principal != null)
                            {
                                if (!lip.Principals.ContainsKey(roleAssignment.Member.LoginName))
                                {
                                    lip.Principals.Add(roleAssignment.Member.LoginName, principal);
                                }
                            }
                        }
                    }
                }
            }

            return lip;
        }

        private Principal GetPrincipal(Web web, string principalInput)
        {
            Principal principal = this.sourceClientContext.Web.SiteGroups.FirstOrDefault(g => g.LoginName.Equals(principalInput, StringComparison.OrdinalIgnoreCase));

            if (principal == null)
            {
                if (principalInput.Contains("#ext#"))
                {
                    principal = web.SiteUsers.FirstOrDefault(u => u.LoginName.Equals(principalInput));

                    if (principal == null)
                    {
                        //Skipping external user...
                    }
                }
                else
                {
                    try
                    {
                        principal = web.EnsureUser(principalInput);
                        web.Context.ExecuteQueryRetry();
                    }
                    catch (Exception ex)
                    {
                        //Failed to EnsureUser
                    }
                }
            }

            return principal;
        }

        private void CopyPageMetadata(PageTransformationInformation pageTransformationInformation, ClientSidePage targetPage, List pagesLibrary)
        {
            var fieldsToCopy = CacheManager.Instance.GetFieldsToCopy(this.sourceClientContext.Web, pagesLibrary);
            if (fieldsToCopy.Count > 0)
            {
                // Load the target page list item
                this.sourceClientContext.Load(targetPage.PageListItem);
                this.sourceClientContext.ExecuteQueryRetry();

                // regular fields
                bool isDirty = false;
                foreach (var fieldToCopy in fieldsToCopy.Where(p => p.FieldType != "TaxonomyFieldTypeMulti" && p.FieldType != "TaxonomyFieldType"))
                {
                    if (pageTransformationInformation.SourcePage[fieldToCopy.FieldName] != null)
                    {
                        targetPage.PageListItem[fieldToCopy.FieldName] = pageTransformationInformation.SourcePage[fieldToCopy.FieldName];
                        isDirty = true;
                    }
                }

                if (isDirty)
                {
                    targetPage.PageListItem.Update();
                    this.sourceClientContext.Load(targetPage.PageListItem);
                    this.sourceClientContext.ExecuteQueryRetry();
                    isDirty = false;
                }

                // taxonomy fields
                foreach (var fieldToCopy in fieldsToCopy.Where(p => p.FieldType == "TaxonomyFieldTypeMulti" || p.FieldType == "TaxonomyFieldType"))
                {
                    switch (fieldToCopy.FieldType)
                    {
                        case "TaxonomyFieldTypeMulti":
                            {
                                var taxFieldBeforeCast = pagesLibrary.Fields.Where(p => p.Id.Equals(fieldToCopy.FieldId)).FirstOrDefault();
                                if (taxFieldBeforeCast != null)
                                {
                                    var taxField = this.sourceClientContext.CastTo<TaxonomyField>(taxFieldBeforeCast);

                                    if (pageTransformationInformation.SourcePage[fieldToCopy.FieldName] != null)
                                    {
                                        if (pageTransformationInformation.SourcePage[fieldToCopy.FieldName] is TaxonomyFieldValueCollection)
                                        {
                                            var valueCollectionToCopy = (pageTransformationInformation.SourcePage[fieldToCopy.FieldName] as TaxonomyFieldValueCollection);
                                            var taxonomyFieldValueArray = valueCollectionToCopy.Select(taxonomyFieldValue => $"-1;#{taxonomyFieldValue.Label}|{taxonomyFieldValue.TermGuid}");
                                            var valueCollection = new TaxonomyFieldValueCollection(this.sourceClientContext, string.Join(";#", taxonomyFieldValueArray), taxField);
                                            taxField.SetFieldValueByValueCollection(targetPage.PageListItem, valueCollection);
                                        }
                                        else if (pageTransformationInformation.SourcePage[fieldToCopy.FieldName] is Dictionary<string, object>)
                                        {
                                            var taxDictionaryList = (pageTransformationInformation.SourcePage[fieldToCopy.FieldName] as Dictionary<string, object>);
                                            var valueCollectionToCopy = taxDictionaryList["_Child_Items_"] as Object[];

                                            List<string> taxonomyFieldValueArray = new List<string>();
                                            for (int i = 0; i < valueCollectionToCopy.Length; i++)
                                            {
                                                var taxDictionary = valueCollectionToCopy[i] as Dictionary<string, object>;
                                                taxonomyFieldValueArray.Add($"-1;#{taxDictionary["Label"].ToString()}|{taxDictionary["TermGuid"].ToString()}");
                                            }
                                            var valueCollection = new TaxonomyFieldValueCollection(this.sourceClientContext, string.Join(";#", taxonomyFieldValueArray), taxField);
                                            taxField.SetFieldValueByValueCollection(targetPage.PageListItem, valueCollection);
                                        }

                                        isDirty = true;
                                    }
                                }
                                break;
                            }
                        case "TaxonomyFieldType":
                            {
                                var taxFieldBeforeCast = pagesLibrary.Fields.Where(p => p.Id.Equals(fieldToCopy.FieldId)).FirstOrDefault();
                                if (taxFieldBeforeCast != null)
                                {
                                    var taxField = this.sourceClientContext.CastTo<TaxonomyField>(taxFieldBeforeCast);
                                    var taxValue = new TaxonomyFieldValue();
                                    if (pageTransformationInformation.SourcePage[fieldToCopy.FieldName] != null)
                                    {
                                        if (pageTransformationInformation.SourcePage[fieldToCopy.FieldName] is TaxonomyFieldValue)
                                        {

                                            taxValue.Label = (pageTransformationInformation.SourcePage[fieldToCopy.FieldName] as TaxonomyFieldValue).Label;
                                            taxValue.TermGuid = (pageTransformationInformation.SourcePage[fieldToCopy.FieldName] as TaxonomyFieldValue).TermGuid;
                                            taxValue.WssId = -1;
                                        }
                                        else if (pageTransformationInformation.SourcePage[fieldToCopy.FieldName] is Dictionary<string, object>)
                                        {
                                            var taxDictionary = (pageTransformationInformation.SourcePage[fieldToCopy.FieldName] as Dictionary<string, object>);
                                            taxValue.Label = taxDictionary["Label"].ToString();
                                            taxValue.TermGuid = taxDictionary["TermGuid"].ToString();
                                            taxValue.WssId = -1;
                                        }
                                        taxField.SetFieldValueByValue(targetPage.PageListItem, taxValue);
                                        isDirty = true;
                                    }
                                }
                                break;
                            }
                    }
                }

                if (isDirty)
                {
                    targetPage.PageListItem.Update();
                    this.sourceClientContext.Load(targetPage.PageListItem);
                    this.sourceClientContext.ExecuteQueryRetry();
                }
            }
        }

        private static void SetPageTitle(PageTransformationInformation pageTransformationInformation, ClientSidePage targetPage)
        {
            if (pageTransformationInformation.SourcePage.FieldExistsAndUsed(Constants.FileLeafRefField))
            {
                string pageTitle = Path.GetFileNameWithoutExtension((pageTransformationInformation.SourcePage[Constants.FileLeafRefField].ToString()));
                if (!string.IsNullOrEmpty(pageTitle))
                {
                    pageTitle = pageTitle.First().ToString().ToUpper() + pageTitle.Substring(1);
                    targetPage.PageTitle = pageTitle;
                }
            }
        }

        private static string GetVersion()
        {
            try
            {
                var coreAssembly = Assembly.GetExecutingAssembly();
                return ((AssemblyFileVersionAttribute)coreAssembly.GetCustomAttribute(typeof(AssemblyFileVersionAttribute))).Version.ToString();
            }
            catch
            {

            }

            return "undefined";
        }

        private Microsoft.SharePoint.Client.File Load(ClientContext cc, PageTransformationInformation pageTransformationInformation, out List pagesLibrary)
        {
            cc.Web.EnsureProperty(w => w.ServerRelativeUrl);

            // Load the pages library and page file (if exists) in one go 
            var listServerRelativeUrl = UrlUtility.Combine(cc.Web.ServerRelativeUrl, "SitePages");
            pagesLibrary = cc.Web.GetList(listServerRelativeUrl);

            if (pageTransformationInformation.CopyPageMetadata)
            {
                cc.Web.Context.Load(pagesLibrary, l => l.DefaultViewUrl, l => l.Id, l => l.BaseTemplate, l => l.OnQuickLaunch, l => l.DefaultViewUrl, l => l.Title, 
                                                  l => l.Hidden, l => l.EffectiveBasePermissions, l => l.RootFolder, l => l.RootFolder.ServerRelativeUrl, 
                                                  l => l.Fields.IncludeWithDefaultProperties(f => f.Id, f => f.Title, f => f.Hidden, f => f.InternalName, f => f.DefaultValue, f => f.Required));
            }
            else
            {
                cc.Web.Context.Load(pagesLibrary, l => l.DefaultViewUrl, l => l.Id, l => l.BaseTemplate, l => l.OnQuickLaunch, l => l.DefaultViewUrl, l => l.Title, 
                                                  l => l.Hidden, l => l.EffectiveBasePermissions, l => l.RootFolder, l => l.RootFolder.ServerRelativeUrl);
            }

            var file = cc.Web.GetFileByServerRelativeUrl($"{listServerRelativeUrl}/{pageTransformationInformation.Folder}{pageTransformationInformation.TargetPageName}");
            cc.Web.Context.Load(file, f => f.Exists, f => f.ListItemAllFields);

            if (pageTransformationInformation.KeepPageSpecificPermissions)
            {
                cc.Load(pageTransformationInformation.SourcePage, p => p.HasUniqueRoleAssignments);
            }

            try
            {
                cc.ExecuteQueryRetry();
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
                throw new ArgumentException($"Site does not have a sitepages library and therefore this page can't be a client side page.");
            }

            if (!file.Exists)
            {
                throw new ArgumentException($"Page {pageTransformationInformation.TargetPageName} does not exist in current web");
            }

            return file;
        }

        private void InitMeasurement()
        {
            try
            {
                if (System.IO.File.Exists(ExecutionLog))
                {
                    System.IO.File.Delete(ExecutionLog);
                }
            }
            catch { }
        }

        private void Start()
        {
            watch = Stopwatch.StartNew();
        }

        private void Stop(string method)
        {
            watch.Stop();
            var elapsedTime = watch.ElapsedMilliseconds;
            System.IO.File.AppendAllText(ExecutionLog, $"{method};{elapsedTime}{Environment.NewLine}");
        }

        /// <summary>
        /// Loads the telemetry and properties for the client object
        /// </summary>
        /// <param name="clientContext"></param>
        private void LoadClientObject(ClientContext clientContext)
        {
            if (clientContext != null)
            {
                clientContext.ClientTag = $"SPDev:PageTransformator";
                // Load all web properties needed further one
                clientContext.Load(clientContext.Web, p => p.Id, p => p.ServerRelativeUrl, p => p.RootFolder.WelcomePage, p => p.Url, p => p.WebTemplate);
                clientContext.Load(clientContext.Site, p => p.RootWeb.ServerRelativeUrl, p => p.Id);
                // Use regular ExecuteQuery as we want to send this custom clienttag
                clientContext.ExecuteQuery();
            }
        }
        #endregion

    }
}
