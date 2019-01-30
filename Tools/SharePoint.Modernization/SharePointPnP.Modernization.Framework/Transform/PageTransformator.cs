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
        private ClientContext clientContext;
        private PageTransformation pageTransformation;
        private string version = "undefined";
        private PageTelemetry pageTelemetry;
        private Stopwatch watch;
        private const string ExecutionLog = "execution.csv";

        #region Construction
        /// <summary>
        /// Creates a page transformator instance
        /// </summary>
        /// <param name="clientContext">ClientContext of the site holding the page</param>
        public PageTransformator(ClientContext clientContext): this(clientContext, "webpartmapping.xml")
        {
        }

        /// <summary>
        /// Creates a page transformator instance
        /// </summary>
        /// <param name="clientContext">ClientContext of the site holding the page</param>
        /// <param name="pageTransformationFile">Used page mapping file</param>
        public PageTransformator(ClientContext clientContext, string pageTransformationFile)
        {

#if DEBUG && MEASURE && MEASURE
            InitMeasurement();
#endif

            this.clientContext = clientContext;
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
        public PageTransformator(ClientContext clientContext, PageTransformation pageTransformationModel)
        {

#if DEBUG && MEASURE
            InitMeasurement();
#endif

            this.clientContext = clientContext;
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
            #endregion

            #region Telemetry
#if DEBUG && MEASURE
            Start();
#endif            
            DateTime transformationStartDateTime = DateTime.Now;
            clientContext.ClientTag = $"SPDev:PageTransformator";
            // Load all web properties needed further one
            clientContext.Load(clientContext.Web, p => p.Id, p => p.ServerRelativeUrl, p => p.RootFolder.WelcomePage, p => p.Url);
            clientContext.Load(clientContext.Site, p => p.RootWeb.ServerRelativeUrl, p => p.Id);
            // Use regular ExecuteQuery as we want to send this custom clienttag
            clientContext.ExecuteQuery();
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
                pageFolder = fileRefFieldValue.Replace($"{clientContext.Web.ServerRelativeUrl}/SitePages", "").Trim();

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

                pageTransformationInformation.TargetPageName = $"{pageTransformationInformation.TargetPagePrefix}{pageTransformationInformation.SourcePage[Constants.FileLeafRefField].ToString()}";
            }

            // Check if page name is free to use
#if DEBUG && MEASURE
            Start();
#endif            
            bool pageExists = false;
            ClientSidePage targetPage = null;
            List pagesLibrary = null;
            Microsoft.SharePoint.Client.File existingFile = null;
            try
            {
                // Just try to load the page in the fastest possible manner, we only want to see if the page exists or not
                existingFile = Load(clientContext, pageTransformationInformation, out pagesLibrary);
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

            targetPage = clientContext.Web.AddClientSidePage($"{pageTransformationInformation.Folder}{pageTransformationInformation.TargetPageName}");
            #endregion

            #region Home page handling
#if DEBUG && MEASURE
            Start();
#endif
            bool replacedByOOBHomePage = false;
            // Check if the transformed page is the web's home page
            if (clientContext.Web.RootFolder.IsPropertyAvailable("WelcomePage") && !string.IsNullOrEmpty(clientContext.Web.RootFolder.WelcomePage))
            {
                var homePageUrl = clientContext.Web.RootFolder.WelcomePage;
                var homepageName = Path.GetFileName(clientContext.Web.RootFolder.WelcomePage);
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
                        Uri host = new Uri(clientContext.Web.Url);

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
            }
            #endregion

            #region Page persisting + permissions
            #region Save the page
#if DEBUG && MEASURE
            Start();
#endif            
            // Persist the client side page
            targetPage.Save($"{pageTransformationInformation.Folder}{pageTransformationInformation.TargetPageName}", existingFile, pagesLibrary);

            // Tag the file with a page modernization version stamp
            try
            {
                string path = pageTransformationInformation.SourcePage[Constants.FileRefField].ToString().Replace(pageTransformationInformation.SourcePage[Constants.FileLeafRefField].ToString(), "");
                var targetPageUrl = $"{path}{pageTransformationInformation.TargetPageName}";
                var targetPageFile = this.clientContext.Web.GetFileByServerRelativeUrl(targetPageUrl);
                this.clientContext.Load(targetPageFile, p => p.Properties);
                targetPageFile.Properties["sharepointpnp_pagemodernization"] = this.version;
                targetPageFile.Update();

                // Try to publish, if publish is not needed then this will return an error that we'll be ignoring
                targetPageFile.Publish("Page modernization initial publish");

                this.clientContext.ExecuteQueryRetry();
            }
            catch (Exception ex)
            {
                // Eat exceptions as this is not critical for the generated page
            }
            
#if DEBUG && MEASURE
            Stop("Persist page");
#endif
            #endregion

            #region Permission handling
            if (pageTransformationInformation.KeepPageSpecificPermissions)
            {
#if DEBUG && MEASURE
                Start();
#endif            
                if (pageTransformationInformation.SourcePage.HasUniqueRoleAssignments)
                {
                    // Copy the unique permissions from source to target
                    // Get the unique permissions
                    this.clientContext.Load(pageTransformationInformation.SourcePage, a => a.RoleAssignments.Include(roleAsg => roleAsg.Member.LoginName,
                        roleAsg => roleAsg.RoleDefinitionBindings.Include(roleDef => roleDef.Name, roleDef => roleDef.Description)));
                    this.clientContext.ExecuteQueryRetry();

                    // Get target page information
                    this.clientContext.Load(targetPage.PageListItem, p => p.HasUniqueRoleAssignments, a => a.RoleAssignments.Include(roleAsg => roleAsg.Member.LoginName,
                        roleAsg => roleAsg.RoleDefinitionBindings.Include(roleDef => roleDef.Name, roleDef => roleDef.Description)));
                    this.clientContext.ExecuteQueryRetry();

                    // Break permission inheritance on the target page if not done yet
                    if (!targetPage.PageListItem.HasUniqueRoleAssignments)
                    {
                        targetPage.PageListItem.BreakRoleInheritance(false, false);
                        this.clientContext.ExecuteQueryRetry();
                    }

                    // Apply new permissions
                    foreach(var roleAssignment in pageTransformationInformation.SourcePage.RoleAssignments)
                    {
                        var principal = this.clientContext.Web.SiteUsers.GetByLoginName(roleAssignment.Member.LoginName);
                        if (principal != null)
                        {
                            var roleDefinitionBindingCollection = new RoleDefinitionBindingCollection(this.clientContext);
                            foreach(var roleDef in roleAssignment.RoleDefinitionBindings)
                            {
                                roleDefinitionBindingCollection.Add(roleDef);
                            }

                            targetPage.PageListItem.RoleAssignments.Add(principal, roleDefinitionBindingCollection);
                        }
                    }
                    this.clientContext.ExecuteQueryRetry();
                }
#if DEBUG && MEASURE
                Stop("Permission handling");
#endif
            }
            #endregion

            #region Page metadata handling
            if (pageTransformationInformation.CopyPageMetadata)
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


            #region Page name switching
            // All went well so far...swap pages if that's needed
            if (pageTransformationInformation.TargetPageTakesSourcePageName)
            {
#if DEBUG && MEASURE
                Start();
#endif            
                //Load the source page
                SwapPages(pageTransformationInformation);
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
        public void SwapPages(PageTransformationInformation pageTransformationInformation)
        {
            var sourcePageUrl = pageTransformationInformation.SourcePage[Constants.FileRefField].ToString();
            var orginalSourcePageName = pageTransformationInformation.SourcePage[Constants.FileLeafRefField].ToString();

            string path = sourcePageUrl.Replace(pageTransformationInformation.SourcePage[Constants.FileLeafRefField].ToString(), "");

            var sourcePage = this.clientContext.Web.GetFileByServerRelativeUrl(sourcePageUrl);
            this.clientContext.Load(sourcePage);
            this.clientContext.ExecuteQueryRetry();

            if (string.IsNullOrEmpty(pageTransformationInformation.SourcePagePrefix))
            {
                pageTransformationInformation.SetDefaultSourcePagePrefix();
            }
            var newSourcePageUrl = $"{pageTransformationInformation.SourcePagePrefix}{pageTransformationInformation.SourcePage[Constants.FileLeafRefField].ToString()}";

            // Rename source page using the sourcepageprefix
            // STEP1: First copy the source page to a new name. We on purpose use CopyTo as we want to avoid that "linked" url's get 
            //        patched up during a MoveTo operation as that would also patch the url's in our new modern page
            sourcePage.CopyTo($"{path}{newSourcePageUrl}", true);
            this.clientContext.ExecuteQueryRetry();

            //Load the created target page
            var targetPageUrl = $"{path}{pageTransformationInformation.TargetPageName}";
            var targetPageFile = this.clientContext.Web.GetFileByServerRelativeUrl(targetPageUrl);
            this.clientContext.Load(targetPageFile);
            this.clientContext.ExecuteQueryRetry();

            // STEP2: Fix possible navigation entries to point to the "copied" source page first
            // Rename the target page to the original source page name
            // CopyTo and MoveTo with option to overwrite first internally delete the file to overwrite, which
            // results in all page navigation nodes pointing to this file to be deleted. Hence let's point these
            // navigation entries first to the copied version of the page we just created
            this.clientContext.Web.Context.Load(this.clientContext.Web, w => w.Navigation.QuickLaunch, w => w.Navigation.TopNavigationBar);
            this.clientContext.Web.Context.ExecuteQueryRetry();

            bool navWasFixed = false;
            IQueryable<NavigationNode> currentNavNodes = null;
            IQueryable<NavigationNode> globalNavNodes = null;
            var currentNavigation = this.clientContext.Web.Navigation.QuickLaunch;
            var globalNavigation = this.clientContext.Web.Navigation.TopNavigationBar;
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
                this.clientContext.ExecuteQueryRetry();
            }

            // STEP3: Now copy the created modern page over the original source page, at this point the new page has the same name as the original page had before transformation
            targetPageFile.CopyTo($"{path}{orginalSourcePageName}", true);
            this.clientContext.ExecuteQueryRetry();

            // STEP4: Finish with restoring the page navigation: update the navlinks to point back the original page name
            if (navWasFixed)
            {
                // Reload the navigation entries as did update them
                this.clientContext.Web.Context.Load(this.clientContext.Web, w => w.Navigation.QuickLaunch, w => w.Navigation.TopNavigationBar);
                this.clientContext.Web.Context.ExecuteQueryRetry();

                currentNavigation = this.clientContext.Web.Navigation.QuickLaunch;
                globalNavigation = this.clientContext.Web.Navigation.TopNavigationBar;
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
                this.clientContext.ExecuteQueryRetry();
            }

            //STEP5: Conclude with deleting the originally created modern page as we did copy that already in step 3
            targetPageFile.DeleteObject();
            this.clientContext.ExecuteQueryRetry();
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
        private void CopyPageMetadata(PageTransformationInformation pageTransformationInformation, ClientSidePage targetPage, List pagesLibrary)
        {
            var fieldsToCopy = CacheManager.Instance.GetFieldsToCopy(this.clientContext.Web, pagesLibrary);
            if (fieldsToCopy.Count > 0)
            {
                // Load the target page list item
                this.clientContext.Load(targetPage.PageListItem);
                this.clientContext.ExecuteQueryRetry();

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
                    this.clientContext.Load(targetPage.PageListItem);
                    this.clientContext.ExecuteQueryRetry();
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
                                    var taxField = this.clientContext.CastTo<TaxonomyField>(taxFieldBeforeCast);

                                    if (pageTransformationInformation.SourcePage[fieldToCopy.FieldName] != null)
                                    {
                                        if (pageTransformationInformation.SourcePage[fieldToCopy.FieldName] is TaxonomyFieldValueCollection)
                                        {
                                            var valueCollectionToCopy = (pageTransformationInformation.SourcePage[fieldToCopy.FieldName] as TaxonomyFieldValueCollection);
                                            var taxonomyFieldValueArray = valueCollectionToCopy.Select(taxonomyFieldValue => $"-1;#{taxonomyFieldValue.Label}|{taxonomyFieldValue.TermGuid}");
                                            var valueCollection = new TaxonomyFieldValueCollection(this.clientContext, string.Join(";#", taxonomyFieldValueArray), taxField);
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
                                            var valueCollection = new TaxonomyFieldValueCollection(this.clientContext, string.Join(";#", taxonomyFieldValueArray), taxField);
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
                                    var taxField = this.clientContext.CastTo<TaxonomyField>(taxFieldBeforeCast);
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
                    this.clientContext.Load(targetPage.PageListItem);
                    this.clientContext.ExecuteQueryRetry();
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
                                                  l => l.Hidden, l => l.RootFolder, l => l.RootFolder.ServerRelativeUrl, 
                                                  l => l.Fields.IncludeWithDefaultProperties(f => f.Id, f => f.Title, f => f.Hidden, f => f.InternalName, f => f.DefaultValue, f => f.Required));
            }
            else
            {
                cc.Web.Context.Load(pagesLibrary, l => l.DefaultViewUrl, l => l.Id, l => l.BaseTemplate, l => l.OnQuickLaunch, l => l.DefaultViewUrl, l => l.Title, 
                                                  l => l.Hidden, l => l.RootFolder, l => l.RootFolder.ServerRelativeUrl);
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
        #endregion

    }
}
