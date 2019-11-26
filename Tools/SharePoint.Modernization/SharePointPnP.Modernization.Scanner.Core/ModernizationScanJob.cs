using System;
using System.Threading;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;
using System.Collections.Concurrent;
using SharePoint.Modernization.Scanner.Core.Results;
using SharePoint.Modernization.Scanner.Core.Analyzers;
using Microsoft.Online.SharePoint.TenantAdministration;
using System.IO;
using System.Linq;
using SharePointPnP.Modernization.Framework;
using SharePoint.Modernization.Scanner.Core.Telemetry;
using SharePoint.Modernization.Scanner.Core.Utilities;
using SharePoint.Modernization.Scanner.Core.Workflow;
using OfficeDevPnP.Core.Framework.Graph;
using Newtonsoft.Json.Linq;
using SharePointPnP.Modernization.Scanner.Core.Analyzers;

namespace SharePoint.Modernization.Scanner.Core
{
    public class ModernizationScanJob: ScanJob
    {
        #region Variables
        private Int32 SitesToScan = 0;
        private string CurrentVersion;
        private string NewVersion;
        private readonly string USDateFormat = "M/d/yyyy";
        private readonly string EuropeDateFormat = "d/M/yyyy";
        private Options options;
        private static volatile bool delveBlogsDone = false;

        internal AppOnlyManager AppOnlyManager;

        public Mode Mode;
        public bool ExportWebPartProperties;
        public bool SkipUsageInformation;
        public bool SkipUserInformation;
        public string DateFormat;
        public bool ExcludeListsOnlyBlockedByOobReasons;
        public string EveryoneExceptExternalUsersClaim = "";
        public readonly string EveryoneClaim = "c:0(.s|true";
        public bool AppOnlyHasFullControl = true;
        public ConcurrentDictionary<string, SiteScanResult> SiteScanResults;
        public ConcurrentDictionary<string, WebScanResult> WebScanResults;
        public ConcurrentDictionary<string, PageScanResult> PageScanResults;
        public ConcurrentDictionary<string, ListScanResult> ListScanResults;
        public Dictionary<string, PublishingSiteScanResult> PublishingSiteScanResults;
        public ConcurrentDictionary<string, PublishingWebScanResult> PublishingWebScanResults;
        public ConcurrentDictionary<string, PublishingPageScanResult> PublishingPageScanResults;
        public ConcurrentDictionary<string, WorkflowScanResult> WorkflowScanResults;
        public ConcurrentDictionary<string, InfoPathScanResult> InfoPathScanResults;
        public ConcurrentDictionary<string, BlogWebScanResult> BlogWebScanResults;
        public ConcurrentDictionary<string, BlogPageScanResult> BlogPageScanResults;
        public List<Guid> TeamifiedSiteCollections;
        public bool TeamifiedSiteCollectionsLoaded = false;
        public Tenant SPOTenant;
        public PageTransformation PageTransformation;
        public ScannerTelemetry ScannerTelemetry;
        #endregion

        #region Construction
        /// <summary>
        /// Instantiate the scanner
        /// </summary>
        /// <param name="options">Options instance</param>
        public ModernizationScanJob(Options options) : base(options, "ModernizationScanner", "1.0")
        {
            this.options = options;

            ExpandSubSites = false;
            Mode = options.Mode;
            ExportWebPartProperties = options.ExportWebPartProperties;
            SkipUsageInformation = options.SkipUsageInformation;
            SkipUserInformation = options.SkipUserInformation;
            ExcludeListsOnlyBlockedByOobReasons = options.ExcludeListsOnlyBlockedByOobReasons;
            CurrentVersion = options.CurrentVersion;
            NewVersion = options.NewVersion;

            // Handle date format setting
            if (options.DateFormat.Equals(USDateFormat) || options.DateFormat.Equals(EuropeDateFormat))
            {
                DateFormat = options.DateFormat;
            }
            else
            {
                // Default to US format in case the provided format is not valid
                DateFormat = USDateFormat;
            }

            this.AppOnlyManager = new AppOnlyManager();

            // Scan results
            this.SiteScanResults = new ConcurrentDictionary<string, SiteScanResult>(options.Threads, 10000);
            this.WebScanResults = new ConcurrentDictionary<string, WebScanResult>(options.Threads, 50000);
            this.ListScanResults = new ConcurrentDictionary<string, ListScanResult>(options.Threads, 100000);
            this.PageScanResults = new ConcurrentDictionary<string, PageScanResult>(options.Threads, 1000000);
            this.WorkflowScanResults = new ConcurrentDictionary<string, WorkflowScanResult>(options.Threads, 100000);
            this.InfoPathScanResults = new ConcurrentDictionary<string, InfoPathScanResult>(options.Threads, 10000);
            this.PublishingSiteScanResults = new Dictionary<string, PublishingSiteScanResult>(500);
            this.PublishingWebScanResults = new ConcurrentDictionary<string, PublishingWebScanResult>(options.Threads, 1000);
            this.PublishingPageScanResults = new ConcurrentDictionary<string, PublishingPageScanResult>(options.Threads, 10000);
            this.BlogWebScanResults = new ConcurrentDictionary<string, BlogWebScanResult>(options.Threads, 50000);
            this.BlogPageScanResults = new ConcurrentDictionary<string, BlogPageScanResult>(options.Threads, 500000);
            this.GeneratedFileStreams = new Dictionary<string, Stream>();
            this.TeamifiedSiteCollections = new List<Guid>();

            // Setup telemetry client
            if (!options.DisableTelemetry)
            {
                this.ScannerTelemetry = new ScannerTelemetry();
                
                // Log scan start event
                if (this.ScannerTelemetry != null)
                {
                    this.ScannerTelemetry.LogScanStart(options);
                }
            }

            VersionWarning();

            this.TimerJobRun += ModernizationScanJob_TimerJobRun;
        }
        #endregion

        #region Scanner implementation
        private void ModernizationScanJob_TimerJobRun(object sender, OfficeDevPnP.Core.Framework.TimerJobs.TimerJobRunEventArgs e)
        {
            // Validate ClientContext objects
            if (e.WebClientContext == null || e.SiteClientContext == null)
            {
                ScanError error = new ScanError()
                {
                    Error = "No valid ClientContext objects",
                    SiteURL = e.Url,
                    SiteColUrl = e.Url
                };
                this.ScanErrors.Push(error);
                Console.WriteLine("Error for site {1}: {0}", "No valid ClientContext objects", e.Url);

                // bail out
                return;
            }

            // Set timeouts 
            e.SiteClientContext.RequestTimeout = Timeout.Infinite;
            e.WebClientContext.RequestTimeout = Timeout.Infinite;
            e.TenantClientContext.RequestTimeout = Timeout.Infinite;

            // thread safe increase of the sites counter
            IncreaseScannedSites();

            try
            {
                // Set the first site collection done flag + perform base bones telemetry
                SetFirstSiteCollectionDone(e.WebClientContext, this.Name);

                // Manually iterate over the content
                IEnumerable<string> expandedSites = e.SiteClientContext.Site.GetAllSubSites();
                bool isFirstSiteInList = true;
                string siteCollectionUrl = "";
                List<Dictionary<string, string>> pageSearchResults = null;
                Dictionary<string, CustomizedPageStatus> masterPageGalleryCustomization = null;

                foreach (string site in expandedSites)
                {
                    try
                    {
                        // thread safe increase of the webs counter
                        IncreaseScannedWebs();

                        // Clone the existing ClientContext for the sub web
                        using (ClientContext ccWeb = e.SiteClientContext.Clone(site))
                        {
                            Console.WriteLine("Processing site {0}...", site);

                            // Allow max server time out, might be needed for sites having a lot of users
                            ccWeb.RequestTimeout = Timeout.Infinite;

                            if (isFirstSiteInList)
                            {
                                // Perf optimization: do one call per site to load all the needed properties
                                var spSite = (ccWeb as ClientContext).Site;
                                ccWeb.Load(spSite, p => p.Url, p => p.GroupId, p => p.Id,
                                                   p => p.RootWeb, p => p.RootWeb.Id,
                                                   p => p.UserCustomActions, // User custom action site level
                                                   p => p.Features // Features site level
                                          );
                                ccWeb.ExecuteQueryRetry();

                                isFirstSiteInList = false;
                            }

                            // Perf optimization: do one call per web to load all the needed properties
                            // Also load the Site RootWeb and Id again as we've a new client context object and this data is needed for the IsSubSite check
                            var spSite2 = (ccWeb as ClientContext).Site;
                            ccWeb.Load(spSite2, p => p.RootWeb, p => p.RootWeb.Id);
                            ccWeb.Load(ccWeb.Web, p => p.Id, p => p.Title, p => p.Url,
                                                  p => p.WebTemplate, p => p.Configuration,
                                                  p => p.MasterUrl, p => p.CustomMasterUrl, // master page check
                                                  p => p.AlternateCssUrl, // Alternate CSS
                                                  p => p.UserCustomActions, // Web user custom actions 
                                                  p => p.Language, p => p.AllProperties, p => p.ServerRelativeUrl, // used in publishing analyzer
                                                  p => p.Features,
                                                  p => p.WelcomePage
                                      );
                            ccWeb.ExecuteQueryRetry();

                            // Split load in multiple batches to minimize timeout exceptions
                            if (!SkipUserInformation)
                            {
                                ccWeb.Load(ccWeb.Web, p => p.SiteUsers, p => p.AssociatedOwnerGroup, p => p.AssociatedMemberGroup, p => p.AssociatedVisitorGroup, // site user and groups
                                                      p => p.HasUniqueRoleAssignments, p => p.RoleAssignments, p => p.SiteGroups, p => p.SiteGroups.Include(s => s.Users) // permission inheritance at web level
                                          );
                                ccWeb.ExecuteQueryRetry();

                                ccWeb.Load(ccWeb.Web, p => p.AssociatedOwnerGroup.Users, // users in the Owners group
                                                      p => p.AssociatedMemberGroup.Users, // users in the Members group
                                                      p => p.AssociatedVisitorGroup.Users // users in the Visitors group
                                          );
                                ccWeb.ExecuteQueryRetry();
                            }

                            // Do things only once per site collection
                            if (string.IsNullOrEmpty(siteCollectionUrl))
                            {
                                // Cross check Url property availability
                                ccWeb.Site.EnsureProperty(s => s.Url);
                                siteCollectionUrl = ccWeb.Site.Url;

                                // Site scan
                                SiteAnalyzer siteAnalyzer = new SiteAnalyzer(site, siteCollectionUrl, this);
                                var siteScanDuration = siteAnalyzer.Analyze(ccWeb);
                                pageSearchResults = siteAnalyzer.PageSearchResults;

                                masterPageGalleryCustomization = new Dictionary<string, CustomizedPageStatus>();
                            }

                            // Web scan
                            WebAnalyzer webAnalyzer = new WebAnalyzer(site, siteCollectionUrl, this, pageSearchResults);
                            webAnalyzer.MasterPageGalleryCustomization = masterPageGalleryCustomization;
                            var webScanDuration = webAnalyzer.Analyze(ccWeb);
                            masterPageGalleryCustomization = webAnalyzer.MasterPageGalleryCustomization;
                        }
                    }
                    catch(Exception ex)
                    {
                        ScanError error = new ScanError()
                        {
                            Error = ex.Message,
                            SiteColUrl = e.Url,
                            SiteURL = site,
                            Field1 = "MainWebLoop",
                            Field2 = ex.StackTrace,
                        };

                        // Send error to telemetry to make scanner better
                        if (this.ScannerTelemetry != null)
                        {
                            this.ScannerTelemetry.LogScanError(ex, error);
                        }

                        this.ScanErrors.Push(error);
                        Console.WriteLine("Error for site {1}: {0}", ex.Message, site);
                    }
                }
            }
            catch (Exception ex)
            {
                ScanError error = new ScanError()
                {
                    Error = ex.Message,
                    SiteColUrl = e.Url,
                    SiteURL = e.Url,
                    Field1 = "MainSiteLoop",
                    Field2 = ex.StackTrace,
                };

                // Send error to telemetry to make scanner better
                if (this.ScannerTelemetry != null)
                {
                    this.ScannerTelemetry.LogScanError(ex, error);
                }

                this.ScanErrors.Push(error);
                Console.WriteLine("Error for site {1}: {0}", ex.Message, e.Url);
            }

            // Output the scanning progress
            try
            {
                TimeSpan ts = DateTime.Now.Subtract(this.StartTime);
                Console.WriteLine($"Thread: {Thread.CurrentThread.ManagedThreadId}. Processed {this.ScannedSites} of {this.SitesToScan} site collections ({Math.Round(((float)this.ScannedSites / (float)this.SitesToScan) * 100)}%). Process running for {ts.Days} days, {ts.Hours} hours, {ts.Minutes} minutes and {ts.Seconds} seconds.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error showing progress: {ex.ToString()}");
            }

        }

        /// <summary>
        /// Grab the number of sites that need to be scanned...will be needed to show progress when we're having a long run
        /// </summary>
        /// <param name="addedSites">List of sites found to scan</param>
        /// <returns>Updated list of sites to scan</returns>
        public override List<string> ResolveAddedSites(List<string> addedSites)
        {            
            try
            {
                this.AppOnlyHasFullControl = this.AppOnlyManager.AppOnlyTokenHasFullControl(this.options, addedSites);
            }
            catch (Exception ex)
            {
                // Oops, not critical...let's continue
            }

            // Determine tenant admin center URL
            string tenantAdmin = "";
            if (!string.IsNullOrEmpty(this.TenantAdminSite))
            {
                tenantAdmin = this.TenantAdminSite;
            }
            else
            {
                if (string.IsNullOrEmpty(this.Tenant))
                {
                    this.Tenant = new Uri(addedSites[0]).DnsSafeHost.Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries)[0];
                }

                tenantAdmin = $"https://{this.Tenant}-admin.sharepoint.com";
            }

            List<string> sites = null;
            if (!this.AppOnlyHasFullControl)
            {
                // We don't have permission to enumerate security information, so turning off that options
                this.SkipUserInformation = true;

                // Query DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS list in tenant admin as fall back scenario
                // Using Tenant Admin API or Search with app-only only works when having full control
                string previousRealm = this.Realm;
                this.Realm = GetRealmFromTargetUrl(new Uri(tenantAdmin));
                using (ClientContext ccAdmin = this.CreateClientContext(tenantAdmin))
                {
                    sites = this.AppOnlyManager.ResolveSitesWithoutFullControl(ccAdmin, addedSites, /*this.ExcludeOD4B*/ true);
                }
                this.Realm = previousRealm;
            }
            else
            {
                sites = base.ResolveAddedSites(addedSites);
            }            
            this.SitesToScan = sites.Count;

            //Perform global initialization tasks, things you only want to do once per run
            if (sites.Count > 0)
            {                
                try
                {
                    using (ClientContext cc = this.CreateClientContext(sites[0]))
                    {
                        // The everyone except external users claim is different per tenant, so grab the correct value
                        this.EveryoneExceptExternalUsersClaim = cc.Web.GetEveryoneExceptExternalUsersClaim();
                    }
                }
                catch(Exception)
                {
                    // Catch exceptions here, typical case is if the used site collection was locked. Do one more try with the root site 
                    var uri = new Uri(sites[0]);
                    using (ClientContext cc = this.CreateClientContext($"{uri.Scheme}://{uri.DnsSafeHost}/"))
                    {
                        // The everyone except external users claim is different per tenant, so grab the correct value
                        this.EveryoneExceptExternalUsersClaim = cc.Web.GetEveryoneExceptExternalUsersClaim();
                    }
                }

                // Handle Delve blog scanning
                if (Options.IncludeBlog(this.options.Mode))
                {
                    if (!delveBlogsDone)
                    {
                        var uri = new Uri(sites[0]);
                        string url = $"{uri.Scheme}://{uri.DnsSafeHost}/";
                        using (ClientContext cc = this.CreateClientContext(url))
                        {
                            DelveBlogAnalyzer delveBlogAnalyzer = new DelveBlogAnalyzer(url, url, this);
                            delveBlogAnalyzer.Analyze(cc);
                            delveBlogsDone = true;
                        }
                    }
                }

                // Create list of Teamified site collections
                try
                {
                    var accessToken = AppOnlyManager.GetGraphAccessToken(this.options);

                    if (!string.IsNullOrEmpty(accessToken))
                    {
                        string getGroupsWithATeamsTeam = $"{GraphHttpClient.MicrosoftGraphBetaBaseUri}groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&select=id,resourceProvisioningOptions";

                        var getGroupResult = GraphHttpClient.MakeGetRequestForString(
                            getGroupsWithATeamsTeam,
                            accessToken: accessToken);

                        JObject groupObject = JObject.Parse(getGroupResult);

                        this.TeamifiedSiteCollectionsLoaded = true;
                        foreach (var item in groupObject["value"])
                        {
                            this.TeamifiedSiteCollections.Add(new Guid(item["id"].ToString()));
                        }
                    }
                }
                catch(Exception ex)
                {
                    TeamifiedSiteCollectionsLoaded = false;
                }

                try
                {
                    if (this.ScannerTelemetry != null)
                    {
                        var uri = new Uri(sites[0]);
                        this.ScannerTelemetry.LoadAADTenantId($"{uri.Scheme}://{uri.DnsSafeHost}");                        
                    }
                }
                catch (Exception ex)
                {
                    // Oops, not critical...let's continue
                }
            }

            // Setup tenant context
            this.Realm = GetRealmFromTargetUrl(new Uri(tenantAdmin));
            using (ClientContext ccAdmin = this.CreateClientContext(tenantAdmin))
            {
                this.SPOTenant = new Tenant(ccAdmin);
            }

            // Load the pagetransformation model that the scanner will use
            this.PageTransformation = new PageTransformationManager().LoadPageTransformationModel();

            // Load the workflow models
            if (Options.IncludeWorkflow(this.Mode) && this.AppOnlyHasFullControl)
            {
                WorkflowManager.Instance.LoadWorkflowDefaultActions();
            }

            return sites;
        }

        /// <summary>
        /// Override of the scanner execute method, needed to output our results
        /// </summary>
        /// <returns>Time when scanning was started</returns>
        public override DateTime Execute()
        {
            // Triggers the run of the scanning...will result in ModernizationScanJob_TimerJobRun being called per site collection
            var start = base.Execute();

            // Telemetry
            if (this.ScannerTelemetry != null)
            {
                this.ScannerTelemetry.LogGroupConnectScan(this.SiteScanResults, this.WebScanResults, this.EveryoneClaim, this.EveryoneExceptExternalUsersClaim);
            }

            // Handle the export of the job specific scanning data
            MemoryStream modernizationSiteScanResults = new MemoryStream();
            string[] outputHeaders = new string[] { "SiteCollectionUrl", "SiteUrl",
                                                    "ReadyForGroupify", "GroupifyBlockers", "GroupifyWarnings", "GroupMode", "PermissionWarnings",
                                                    "ModernHomePage", "ModernUIWarnings",
                                                    "WebTemplate", "Office365GroupId", "HasTeamsTeam", "MasterPage", "AlternateCSS", "UserCustomActions",
                                                    "SubSites", "SubSitesWithBrokenPermissionInheritance", "ModernPageWebFeatureDisabled", "ModernPageFeatureWasEnabledBySPO",
                                                    "ModernListSiteBlockingFeatureEnabled", "ModernListWebBlockingFeatureEnabled", "SitePublishingFeatureEnabled", "WebPublishingFeatureEnabled",
                                                    "ViewsRecent", "ViewsRecentUniqueUsers", "ViewsLifeTime", "ViewsLifeTimeUniqueUsers", "SiteId",
                                                    "Everyone(ExceptExternalUsers)Claim", "UsesADGroups", "ExternalSharing",
                                                    "Admins", "AdminContainsEveryone(ExceptExternalUsers)Claim", "AdminContainsADGroups",
                                                    "Owners", "OwnersContainsEveryone(ExceptExternalUsers)Claim", "OwnersContainsADGroups",
                                                    "Members", "MembersContainsEveryone(ExceptExternalUsers)Claim", "MembersContainsADGroups",
                                                    "Visitors", "VisitorsContainsEveryone(ExceptExternalUsers)Claim", "VisitorsContainsADGroups"
                                                  };

            StreamWriter outStream = new StreamWriter(modernizationSiteScanResults);
            outStream.Write(string.Format("{0}\r\n", string.Join(this.Separator, outputHeaders)));
            foreach (var item in this.SiteScanResults)
            {
                var groupifyBlockers = item.Value.GroupifyBlockers();
                var groupifyWarnings = item.Value.GroupifyWarnings(this.EveryoneClaim, this.EveryoneExceptExternalUsersClaim);
                var modernWarnings = item.Value.ModernWarnings();
                var groupSecurity = item.Value.PermissionModel(this.EveryoneClaim, this.EveryoneExceptExternalUsersClaim);

                outStream.Write(string.Format("{0}\r\n", string.Join(this.Separator, ToCsv(item.Value.SiteColUrl), ToCsv(item.Value.SiteURL),
                                                                                       (groupifyBlockers.Count > 0 ? "FALSE" : "TRUE"), ToCsv(SiteScanResult.FormatList(groupifyBlockers)), ToCsv(SiteScanResult.FormatList(groupifyWarnings)), ToCsv(groupSecurity.Item1), ToCsv(SiteScanResult.FormatList(groupSecurity.Item2)),
                                                                                       item.Value.ModernHomePage, ToCsv(SiteScanResult.FormatList(modernWarnings)),
                                                                                       ToCsv(item.Value.WebTemplate), ToCsv(item.Value.Office365GroupId != Guid.Empty ? item.Value.Office365GroupId.ToString() : ""), item.Value.IsTeamified(), item.Value.MasterPage, item.Value.AlternateCSS, ((item.Value.SiteUserCustomActions != null && item.Value.SiteUserCustomActions.Count > 0) || (item.Value.WebUserCustomActions != null && item.Value.WebUserCustomActions.Count > 0)),
                                                                                       item.Value.SubSites, item.Value.SubSitesWithBrokenPermissionInheritance, item.Value.ModernPageWebFeatureDisabled, item.Value.ModernPageFeatureWasEnabledBySPO,
                                                                                       item.Value.ModernListSiteBlockingFeatureEnabled, item.Value.ModernListWebBlockingFeatureEnabled, item.Value.SitePublishingFeatureEnabled, item.Value.WebPublishingFeatureEnabled,
                                                                                       (SkipUsageInformation ? 0 : item.Value.ViewsRecent), (SkipUsageInformation ? 0 : item.Value.ViewsRecentUniqueUsers), (SkipUsageInformation ? 0 : item.Value.ViewsLifeTime), (SkipUsageInformation ? 0 : item.Value.ViewsLifeTimeUniqueUsers), ToCsv(item.Value.SiteId),
                                                                                       item.Value.EveryoneClaimsGranted, item.Value.ContainsADGroup(), ToCsv(item.Value.SharingCapabilities),
                                                                                       ToCsv(SiteScanResult.FormatUserList(item.Value.Admins, this.EveryoneClaim, this.EveryoneExceptExternalUsersClaim)), item.Value.HasClaim(item.Value.Admins, this.EveryoneClaim, this.EveryoneExceptExternalUsersClaim), item.Value.ContainsADGroup(item.Value.Admins),
                                                                                       ToCsv(SiteScanResult.FormatUserList(item.Value.Owners, this.EveryoneClaim, this.EveryoneExceptExternalUsersClaim)), item.Value.HasClaim(item.Value.Owners, this.EveryoneClaim, this.EveryoneExceptExternalUsersClaim), item.Value.ContainsADGroup(item.Value.Owners),
                                                                                       ToCsv(SiteScanResult.FormatUserList(item.Value.Members, this.EveryoneClaim, this.EveryoneExceptExternalUsersClaim)), item.Value.HasClaim(item.Value.Members, this.EveryoneClaim, this.EveryoneExceptExternalUsersClaim), item.Value.ContainsADGroup(item.Value.Members),
                                                                                       ToCsv(SiteScanResult.FormatUserList(item.Value.Visitors, this.EveryoneClaim, this.EveryoneExceptExternalUsersClaim)), item.Value.HasClaim(item.Value.Visitors, this.EveryoneClaim, this.EveryoneExceptExternalUsersClaim), item.Value.ContainsADGroup(item.Value.Visitors)
                                                )));
            }
            outStream.Flush();
            this.GeneratedFileStreams.Add("ModernizationSiteScanResults.csv", modernizationSiteScanResults);

            MemoryStream modernizationWebScanResults = new MemoryStream();
            outputHeaders = new string[] { "SiteCollectionUrl", "SiteUrl",
                                           "WebTemplate", "BrokenPermissionInheritance", "ModernPageWebFeatureDisabled", "ModernPageFeatureWasEnabledBySPO", "WebPublishingFeatureEnabled",
                                           "MasterPage", "CustomMasterPage", "AlternateCSS", "UserCustomActions",
                                           "Everyone(ExceptExternalUsers)Claim",
                                           "UniqueOwners",
                                           "UniqueMembers",
                                           "UniqueVisitors"
                                         };
            outStream = new StreamWriter(modernizationWebScanResults);
            outStream.Write(string.Format("{0}\r\n", string.Join(this.Separator, outputHeaders)));

            foreach (var item in this.WebScanResults)
            {
                outStream.Write(string.Format("{0}\r\n", string.Join(this.Separator, ToCsv(item.Value.SiteColUrl), ToCsv(item.Value.SiteURL),
                                                                                       ToCsv(item.Value.WebTemplate), item.Value.BrokenPermissionInheritance, item.Value.ModernPageWebFeatureDisabled, item.Value.ModernPageFeatureWasEnabledBySPO, item.Value.WebPublishingFeatureEnabled,
                                                                                       ToCsv(item.Value.MasterPage), ToCsv(item.Value.CustomMasterPage), ToCsv(item.Value.AlternateCSS), (item.Value.WebUserCustomActions.Count > 0),
                                                                                       item.Value.EveryoneClaimsGranted,
                                                                                       ToCsv(SiteScanResult.FormatUserList(item.Value.Owners, this.EveryoneClaim, this.EveryoneExceptExternalUsersClaim)),
                                                                                       ToCsv(SiteScanResult.FormatUserList(item.Value.Members, this.EveryoneClaim, this.EveryoneExceptExternalUsersClaim)),
                                                                                       ToCsv(SiteScanResult.FormatUserList(item.Value.Visitors, this.EveryoneClaim, this.EveryoneExceptExternalUsersClaim))
                                                )));
            }

            outStream.Flush();
            this.GeneratedFileStreams.Add("ModernizationWebScanResults.csv", modernizationWebScanResults);

            MemoryStream modernizationUserCustomActionScanResults = new MemoryStream();
            outputHeaders = new string[] { "SiteCollectionUrl", "SiteUrl",
                                           "Title", "Name", "Location", "RegistrationType", "RegistrationId", "Reason", "CommandAction", "ScriptBlock", "ScriptSrc"
                                         };

            outStream = new StreamWriter(modernizationUserCustomActionScanResults);
            outStream.Write(string.Format("{0}\r\n", string.Join(this.Separator, outputHeaders)));

            foreach (var item in this.SiteScanResults)
            {
                if (item.Value.SiteUserCustomActions == null || item.Value.SiteUserCustomActions.Count == 0)
                {
                    continue;
                }

                foreach (var uca in item.Value.SiteUserCustomActions)
                {
                    outStream.Write(string.Format("{0}\r\n", string.Join(this.Separator, ToCsv(item.Value.SiteColUrl), ToCsv(item.Value.SiteURL),
                                                                                           ToCsv(uca.Title), ToCsv(uca.Name), ToCsv(uca.Location), uca.RegistrationType, ToCsv(uca.RegistrationId), ToCsv(uca.Problem), ToCsv(uca.CommandAction), ToCsv(uca.ScriptBlock), ToCsv(uca.ScriptSrc)
                                                     )));
                }
            }
            foreach (var item in this.WebScanResults)
            {
                if (item.Value.WebUserCustomActions == null || item.Value.WebUserCustomActions.Count == 0)
                {
                    continue;
                }

                foreach (var uca in item.Value.WebUserCustomActions)
                {
                    outStream.Write(string.Format("{0}\r\n", string.Join(this.Separator, ToCsv(item.Value.SiteColUrl), ToCsv(item.Value.SiteURL),
                                                                                           ToCsv(uca.Title), ToCsv(uca.Name), ToCsv(uca.Location), uca.RegistrationType, ToCsv(uca.RegistrationId), ToCsv(uca.Problem), ToCsv(uca.CommandAction), ToCsv(uca.ScriptBlock), ToCsv(uca.ScriptSrc)
                                                     )));
                }
            }

            outStream.Flush();
            this.GeneratedFileStreams.Add("ModernizationUserCustomActionScanResults.csv", modernizationUserCustomActionScanResults);

            if (Options.IncludeLists(this.Mode))
            {
                // Telemetry
                if (this.ScannerTelemetry != null)
                {
                    this.ScannerTelemetry.LogListScan(this.ScannedSites, this.ScannedWebs, this.ListScanResults, this.ScannedLists);
                }

                MemoryStream modernizationListScanResults = new MemoryStream();
                outStream = new StreamWriter(modernizationListScanResults);

                outputHeaders = new string[] { "Url", "Site Url", "Site Collection Url", "List Title", "Only blocked by OOB reasons",
                                               "Blocked at site level", "Blocked at web level", "Blocked at list level", "List page render type", "List experience", "Blocked by not being able to load Page", "Blocked by not being able to load page exception",
                                               "Blocked by managed metadata navigation", "Blocked by view type", "View type", "Blocked by list base template", "List base template",
                                               "Blocked by zero or multiple web parts", "Blocked by JSLink", "JSLink", "Blocked by XslLink", "XslLink", "Blocked by Xsl",
                                               "Blocked by JSLink field", "JSLink fields", "Blocked by business data field", "Business data fields", "Blocked by task outcome field", "Task outcome fields",
                                               "Blocked by publishingField", "Publishing fields", "Blocked by geo location field", "Geo location fields", "Blocked by list custom action", "List custom actions"  };

                outStream.Write(string.Format("{0}\r\n", string.Join(this.Separator, outputHeaders)));
                foreach (var list in this.ListScanResults)
                {

                    outStream.Write(string.Format("{0}\r\n", string.Join(this.Separator, ToCsv(list.Key.Substring(36)), ToCsv(list.Value.SiteURL), ToCsv(list.Value.SiteColUrl), ToCsv(list.Value.ListTitle), list.Value.OnlyBlockedByOOBReasons,
                                                           list.Value.BlockedAtSiteLevel, list.Value.BlockedAtWebLevel, list.Value.BlockedAtListLevel, list.Value.PageRenderType, list.Value.ListExperience, list.Value.BlockedByNotBeingAbleToLoadPage, ToCsv(list.Value.BlockedByNotBeingAbleToLoadPageException),
                                                           list.Value.XsltViewWebPartCompatibility.BlockedByManagedMetadataNavFeature, list.Value.XsltViewWebPartCompatibility.BlockedByViewType, ToCsv(list.Value.XsltViewWebPartCompatibility.ViewType), list.Value.XsltViewWebPartCompatibility.BlockedByListBaseTemplate, list.Value.XsltViewWebPartCompatibility.ListBaseTemplate,
                                                           list.Value.BlockedByZeroOrMultipleWebParts, list.Value.XsltViewWebPartCompatibility.BlockedByJSLink, ToCsv(list.Value.XsltViewWebPartCompatibility.JSLink), list.Value.XsltViewWebPartCompatibility.BlockedByXslLink, ToCsv(list.Value.XsltViewWebPartCompatibility.XslLink), list.Value.XsltViewWebPartCompatibility.BlockedByXsl,
                                                           list.Value.XsltViewWebPartCompatibility.BlockedByJSLinkField, ToCsv(list.Value.XsltViewWebPartCompatibility.JSLinkFields), list.Value.XsltViewWebPartCompatibility.BlockedByBusinessDataField, ToCsv(list.Value.XsltViewWebPartCompatibility.BusinessDataFields), list.Value.XsltViewWebPartCompatibility.BlockedByTaskOutcomeField, ToCsv(list.Value.XsltViewWebPartCompatibility.TaskOutcomeFields),
                                                           list.Value.XsltViewWebPartCompatibility.BlockedByPublishingField, ToCsv(list.Value.XsltViewWebPartCompatibility.PublishingFields), list.Value.XsltViewWebPartCompatibility.BlockedByGeoLocationField, ToCsv(list.Value.XsltViewWebPartCompatibility.GeoLocationFields), list.Value.XsltViewWebPartCompatibility.BlockedByListCustomAction, ToCsv(list.Value.XsltViewWebPartCompatibility.ListCustomActions)
                                                    )));
                }
                outStream.Flush();
                this.GeneratedFileStreams.Add("ModernizationListScanResults.csv", modernizationListScanResults);

                // Analyze the lists that and export the site collections that use classic customizations
                var sitesWithCodeCustomizations = ListAnalyzer.GenerateSitesWithCodeCustomizationsResults(this.ListScanResults);
                MemoryStream sitesWithCustomizations = new MemoryStream();
                outStream = new StreamWriter(sitesWithCustomizations);

                foreach (var siteWithCustomizations in sitesWithCodeCustomizations)
                {
                    outStream.Write(string.Format("{0}\r\n", string.Join(this.Separator, ToCsv(siteWithCustomizations))));
                }
                outStream.Flush();
                this.GeneratedFileStreams.Add("SitesWithCustomizations.csv", sitesWithCustomizations);

            }

            if (Options.IncludePage(this.Mode))
            {
                // Telemetry
                if (this.ScannerTelemetry != null)
                {
                    this.ScannerTelemetry.LogPageScan(this.ScannedSites, this.ScannedWebs, this.PageScanResults, this.PageTransformation);
                }

                MemoryStream pageScanResults = new MemoryStream();
                outStream = new StreamWriter(pageScanResults);
                outputHeaders = new string[] { "SiteCollectionUrl", "SiteUrl", "PageUrl", "Library", "HomePage", "Uncustomized STS#0 home page",
                                           "Type", "Layout", "Mapping %", "Unmapped web parts", "ModifiedBy", "ModifiedAt",
                                           "ViewsRecent", "ViewsRecentUniqueUsers", "ViewsLifeTime", "ViewsLifeTimeUniqueUsers"};

                string header1 = string.Join(this.Separator, outputHeaders);
                string header2 = "";
                for (int i = 1; i <= 30; i++)
                {
                    if (ExportWebPartProperties)
                    {
                        header2 = header2 + $"{this.Separator}WPType{i}{this.Separator}WPTitle{i}{this.Separator}WPData{i}";
                    }
                    else
                    {
                        header2 = header2 + $"{this.Separator}WPType{i}{this.Separator}WPTitle{i}";
                    }
                }

                List<string> UniqueWebParts = new List<string>();

                outStream.Write(string.Format("{0}\r\n", header1 + header2));
                foreach (var item in this.PageScanResults)
                {
                    var part1 = string.Join(this.Separator, ToCsv(item.Value.SiteColUrl), ToCsv(item.Value.SiteURL), ToCsv(item.Value.PageUrl), ToCsv(item.Value.Library), item.Value.HomePage, item.Value.UncustomizedHomePage,
                                                            ToCsv(item.Value.PageType), ToCsv(item.Value.Layout), "{MappingPercentage}", "{UnmappedWebParts}", ToCsv(item.Value.ModifiedBy), ToDateString(item.Value.ModifiedAt, this.DateFormat),
                                                            (SkipUsageInformation ? 0 : item.Value.ViewsRecent), (SkipUsageInformation ? 0 : item.Value.ViewsRecentUniqueUsers), (SkipUsageInformation ? 0 : item.Value.ViewsLifeTime), (SkipUsageInformation ? 0 : item.Value.ViewsLifeTimeUniqueUsers));

                    string part2 = "";
                    if (item.Value.WebParts != null)
                    {
                        int webPartsOnPage = item.Value.WebParts.Count();
                        int webPartsOnPageMapped = 0;
                        List<string> nonMappedWebParts = new List<string>();
                        foreach (var webPart in item.Value.WebParts.OrderBy(p => p.Row).ThenBy(p => p.Column).ThenBy(p => p.Order))
                        {
                            var found = this.PageTransformation.WebParts.Where(p => p.Type.Equals(webPart.Type, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                            if (found != null && found.Mappings != null)
                            {
                                webPartsOnPageMapped++;
                            }
                            else
                            {
                                var t = webPart.Type.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries)[0];
                                if (!nonMappedWebParts.Contains(t))
                                {
                                    nonMappedWebParts.Add(t);
                                }
                            }

                            if (ExportWebPartProperties)
                            {
                                part2 = part2 + $"{this.Separator}{ToCsv(webPart.TypeShort())}{this.Separator}{ToCsv(webPart.Title)}{this.Separator}{ToCsv(webPart.Json())}";
                            }
                            else
                            {
                                part2 = part2 + $"{this.Separator}{ToCsv(webPart.TypeShort())}{this.Separator}{ToCsv(webPart.Title)}";
                            }

                            if (!UniqueWebParts.Contains(webPart.Type))
                            {
                                UniqueWebParts.Add(webPart.Type);
                            }
                        }
                        part1 = part1.Replace("{MappingPercentage}", webPartsOnPage == 0 ? "100" : String.Format("{0:0}", (((double)webPartsOnPageMapped / (double)webPartsOnPage) * 100))).Replace("{UnmappedWebParts}", SiteScanResult.FormatList(nonMappedWebParts));
                    }
                    else
                    {
                        part1 = part1.Replace("{MappingPercentage}", "").Replace("{UnmappedWebParts}", "");
                    }

                    outStream.Write(string.Format("{0}\r\n", part1 + (!string.IsNullOrEmpty(part2) ? part2 : "")));
                }

                outStream.Flush();
                this.GeneratedFileStreams.Add("PageScanResults.csv", pageScanResults);

                MemoryStream uniqueWebParts = new MemoryStream();
                outStream = new StreamWriter(uniqueWebParts);
                outStream.Write(string.Format("{0}\r\n", $"Type{this.Separator}InMappingFile"));
                foreach (var type in UniqueWebParts)
                {
                    var found = this.PageTransformation.WebParts.Where(p => p.Type.Equals(type, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                    outStream.Write(string.Format("{0}\r\n", $"{ToCsv(type)}{this.Separator}{found != null}"));
                }
                outStream.Flush();
                this.GeneratedFileStreams.Add("UniqueWebParts.csv", uniqueWebParts);
            }

            if (Options.IncludePublishing(this.Mode))
            {
                // "Calculate" publishing site results based upon the web/page level data we retrieved
                this.PublishingSiteScanResults = PublishingAnalyzer.GeneratePublishingSiteResults(this.Mode, this.PublishingWebScanResults, this.PublishingPageScanResults);

                // Telemetry
                if (this.ScannerTelemetry != null)
                {
                    this.ScannerTelemetry.LogPublishingScan(this.PublishingSiteScanResults, this.PublishingWebScanResults, this.PublishingPageScanResults, this.PageTransformation);
                }

                // Export the site publishing data
                MemoryStream modernizationPublishingSiteScanResults = new MemoryStream();
                outputHeaders = new string[] { "SiteCollectionUrl", "NumberOfWebs", "NumberOfPages",
                                               "UsedSiteMasterPages", "UsedSystemMasterPages",
                                               "UsedPageLayouts", "LastPageUpdateDate"
                                             };
                outStream = new StreamWriter(modernizationPublishingSiteScanResults);

                outStream.Write(string.Format("{0}\r\n", string.Join(this.Separator, outputHeaders)));
                if (PublishingSiteScanResults != null)
                {
                    foreach (var item in this.PublishingSiteScanResults)
                    {
                        outStream.Write(string.Format("{0}\r\n", string.Join(this.Separator, ToCsv(item.Value.SiteColUrl), item.Value.NumberOfWebs, item.Value.NumberOfPages,
                                                                                               ToCsv(PublishingPageScanResult.FormatList(item.Value.UsedSiteMasterPages)), ToCsv(PublishingPageScanResult.FormatList(item.Value.UsedSystemMasterPages)),
                                                                                               ToCsv(PublishingPageScanResult.FormatList(item.Value.UsedPageLayouts)), item.Value.LastPageUpdateDate.HasValue ? item.Value.LastPageUpdateDate.ToString() : ""
                                                       )));
                    }
                }
                outStream.Flush();
                this.GeneratedFileStreams.Add("ModernizationPublishingSiteScanResults.csv", modernizationPublishingSiteScanResults);

                // Export the web publishing data
                MemoryStream modernizationPublishingWebScanResults = new MemoryStream();

                outputHeaders = new string[] { "SiteCollectionUrl", "SiteUrl", "WebRelativeUrl", "SiteCollectionComplexity",
                                               "WebTemplate", "Level", "PageCount", "Language", "VariationLabels", "VariationSourceLabel",
                                               "SiteMasterPage", "SystemMasterPage", "AlternateCSS", "HasIncompatibleUserCustomActions",
                                               "AllowedPageLayouts", "PageLayoutsConfiguration", "DefaultPageLayout",
                                               "GlobalNavigationType", "GlobalStructuralNavigationShowSubSites", "GlobalStructuralNavigationShowPages","GlobalStructuralNavigationShowSiblings","GlobalStructuralNavigationMaxCount","GlobalManagedNavigationTermSetId",
                                               "CurrentNavigationType","CurrentStructuralNavigationShowSubSites","CurrentStructuralNavigationShowPages","CurrentStructuralNavigationShowSiblings","CurrentStructuralNavigationMaxCount","CurrentManagedNavigationTermSetId",
                                               "ManagedNavigationAddNewPages", "ManagedNavigationCreateFriendlyUrls",
                                               "LibraryItemScheduling","LibraryEnableModeration","LibraryEnableVersioning","LibraryEnableMinorVersions","LibraryApprovalWorkflowDefined",
                                               "BrokenPermissionInheritance",
                                               "Admins",
                                               "Owners"
                                             };
                outStream = new StreamWriter(modernizationPublishingWebScanResults);

                outStream.Write(string.Format("{0}\r\n", string.Join(this.Separator, outputHeaders)));
                foreach (var item in this.PublishingWebScanResults)
                {
                    outStream.Write(string.Format("{0}\r\n", string.Join(this.Separator, ToCsv(item.Value.SiteColUrl), ToCsv(item.Value.SiteURL), ToCsv(item.Value.WebRelativeUrl), ToCsv(item.Value.SiteClassification),
                                                                                           ToCsv(item.Value.WebTemplate), item.Value.Level.ToString(), item.Value.PageCount.ToString(), item.Value.Language.ToString(), ToCsv(item.Value.VariationLabels), ToCsv(item.Value.VariationSourceLabel),
                                                                                           ToCsv(item.Value.SiteMasterPage), ToCsv(item.Value.SystemMasterPage), ToCsv(item.Value.AlternateCSS), (item.Value.UserCustomActions != null && item.Value.UserCustomActions.Count > 0),
                                                                                           ToCsv(item.Value.AllowedPageLayouts), ToCsv(item.Value.PageLayoutsConfiguration), ToCsv(item.Value.DefaultPageLayout),
                                                                                           ToCsv(item.Value.GlobalNavigationType), item.Value.GlobalStructuralNavigationShowSubSites.HasValue ? item.Value.GlobalStructuralNavigationShowSubSites.Value.ToString() : "", item.Value.GlobalStructuralNavigationShowPages.HasValue ? item.Value.GlobalStructuralNavigationShowPages.Value.ToString() : "", item.Value.GlobalStructuralNavigationShowSiblings.HasValue ? item.Value.GlobalStructuralNavigationShowSiblings.Value.ToString() : "", item.Value.GlobalStructuralNavigationMaxCount.HasValue ? item.Value.GlobalStructuralNavigationMaxCount.Value.ToString() : "", ToCsv(item.Value.GlobalManagedNavigationTermSetId),
                                                                                           ToCsv(item.Value.CurrentNavigationType), item.Value.CurrentStructuralNavigationShowSubSites.HasValue ? item.Value.CurrentStructuralNavigationShowSubSites.Value.ToString() : "", item.Value.CurrentStructuralNavigationShowPages.HasValue ? item.Value.CurrentStructuralNavigationShowPages.Value.ToString() : "", item.Value.CurrentStructuralNavigationShowSiblings.HasValue ? item.Value.CurrentStructuralNavigationShowSiblings.Value.ToString() : "", item.Value.CurrentStructuralNavigationMaxCount.HasValue ? item.Value.CurrentStructuralNavigationMaxCount.Value.ToString() : "", ToCsv(item.Value.CurrentManagedNavigationTermSetId),
                                                                                           item.Value.ManagedNavigationAddNewPages.HasValue ? item.Value.ManagedNavigationAddNewPages.ToString() : "", item.Value.ManagedNavigationCreateFriendlyUrls.HasValue ? item.Value.ManagedNavigationCreateFriendlyUrls.ToString() : "",
                                                                                           item.Value.LibraryItemScheduling.ToString(), item.Value.LibraryEnableModeration.ToString(), item.Value.LibraryEnableVersioning.ToString(), item.Value.LibraryEnableMinorVersions.ToString(), item.Value.LibraryApprovalWorkflowDefined.ToString(),
                                                                                           item.Value.BrokenPermissionInheritance.ToString(),
                                                                                           ToCsv(SiteScanResult.FormatUserList(item.Value.Admins, this.EveryoneClaim, this.EveryoneExceptExternalUsersClaim)),
                                                                                           ToCsv(SiteScanResult.FormatUserList(item.Value.Owners, this.EveryoneClaim, this.EveryoneExceptExternalUsersClaim))
                                                    )));
                }
                outStream.Flush();
                this.GeneratedFileStreams.Add("ModernizationPublishingWebScanResults.csv", modernizationPublishingWebScanResults);

                if (Options.IncludePublishingWithPages(this.Mode))
                {
                    // Export the page publishing data
                    MemoryStream modernizationPublishingPageScanResults = new MemoryStream();
                    outputHeaders = new string[] { "SiteCollectionUrl", "SiteUrl", "WebRelativeUrl", "PageRelativeUrl", "PageName",
                                                   "ContentType", "ContentTypeId", "PageLayout", "PageLayoutFile", "PageLayoutWasCustomized",
                                                   "GlobalAudiences", "SecurityGroupAudiences", "SharePointGroupAudiences",
                                                   "ModifiedAt", "ModifiedBy", "Mapping %", "Unmapped web parts"
                                                 };

                    string header1 = string.Join(this.Separator, outputHeaders);
                    string header2 = "";
                    for (int i = 1; i <= 20; i++)
                    {
                        if (ExportWebPartProperties)
                        {
                            header2 = header2 + $"{this.Separator}WPType{i}{this.Separator}WPTitle{i}{this.Separator}WPData{i}";
                        }
                        else
                        {
                            header2 = header2 + $"{this.Separator}WPType{i}{this.Separator}WPTitle{i}";
                        }
                    }

                    outStream = new StreamWriter(modernizationPublishingPageScanResults);

                    outStream.Write(string.Format("{0}\r\n", header1 + header2));
                    foreach (var item in this.PublishingPageScanResults)
                    {
                        var part1 = string.Join(this.Separator, ToCsv(item.Value.SiteColUrl), ToCsv(item.Value.SiteURL), ToCsv(item.Value.WebRelativeUrl), ToCsv(item.Value.PageRelativeUrl), ToCsv(item.Value.PageName),
                                                                ToCsv(item.Value.ContentType), ToCsv(item.Value.ContentTypeId), ToCsv(item.Value.PageLayout), ToCsv(item.Value.PageLayoutFile), item.Value.PageLayoutWasCustomized,
                                                                ToCsv(PublishingPageScanResult.FormatList(item.Value.GlobalAudiences)), ToCsv(PublishingPageScanResult.FormatList(item.Value.SecurityGroupAudiences, "|")), ToCsv(PublishingPageScanResult.FormatList(item.Value.SharePointGroupAudiences)),
                                                                ToDateString(item.Value.ModifiedAt, this.DateFormat), ToCsv(item.Value.ModifiedBy), "{MappingPercentage}", "{UnmappedWebParts}"
                            );

                        string part2 = "";
                        if (item.Value.WebParts != null)
                        {
                            int webPartsOnPage = item.Value.WebParts.Count();
                            int webPartsOnPageMapped = 0;
                            List<string> nonMappedWebParts = new List<string>();
                            foreach (var webPart in item.Value.WebParts.OrderBy(p => p.Row).ThenBy(p => p.Column).ThenBy(p => p.Order))
                            {
                                var found = this.PageTransformation.WebParts.Where(p => p.Type.Equals(webPart.Type, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                                if (found != null && found.Mappings != null)
                                {
                                    webPartsOnPageMapped++;
                                }
                                else
                                {
                                    var t = webPart.Type.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries)[0];
                                    if (!nonMappedWebParts.Contains(t))
                                    {
                                        nonMappedWebParts.Add(t);
                                    }
                                }

                                if (ExportWebPartProperties)
                                {
                                    part2 = part2 + $"{this.Separator}{ToCsv(webPart.TypeShort())}{this.Separator}{ToCsv(webPart.Title)}{this.Separator}{ToCsv(webPart.Json())}";
                                }
                                else
                                {
                                    part2 = part2 + $"{this.Separator}{ToCsv(webPart.TypeShort())}{this.Separator}{ToCsv(webPart.Title)}";
                                }
                            }
                            part1 = part1.Replace("{MappingPercentage}", webPartsOnPage == 0 ? "100" : String.Format("{0:0}", (((double)webPartsOnPageMapped / (double)webPartsOnPage) * 100))).Replace("{UnmappedWebParts}", SiteScanResult.FormatList(nonMappedWebParts));
                        }
                        else
                        {
                            part1 = part1.Replace("{MappingPercentage}", "").Replace("{UnmappedWebParts}", "");
                        }

                        outStream.Write(string.Format("{0}\r\n", part1 + (!string.IsNullOrEmpty(part2) ? part2 : "")));
                    }
                    outStream.Flush();
                    this.GeneratedFileStreams.Add("ModernizationPublishingPageScanResults.csv", modernizationPublishingPageScanResults);
                }
            }

            if (Options.IncludeWorkflow(this.Mode))
            {
                // Telemetry
                if (this.ScannerTelemetry != null)
                {
                    this.ScannerTelemetry.LogWorkflowScan(this.WorkflowScanResults);
                }

                MemoryStream modernizationWorkflowScanResults = new MemoryStream();
                outputHeaders = new string[] { "Site Url", "Site Collection Url", "Definition Name", "Migration to Flow recommended", "Version", "Scope", "Has subscriptions", "Enabled", "Is OOB",
                                               "List Title", "List Url", "List Id", "ContentType Name", "ContentType Id",
                                               "Restricted To", "Definition description", "Definition Id", "Subscription Name", "Subscription Id",
                                               "Definition Changed On", "Subscription Changed On",
                                               "Action Count", "Used Actions", "Used Triggers", "Flow upgradability", "Incompatible Action Count", "Incompatible Actions",
                                               "Change Year", "Change Quarter", "Change Month" };

                outStream = new StreamWriter(modernizationWorkflowScanResults);


                outStream.Write(string.Format("{0}\r\n", string.Join(this.Separator, outputHeaders)));
                foreach (var workflow in this.WorkflowScanResults)
                {

                    outStream.Write(string.Format("{0}\r\n", string.Join(this.Separator, ToCsv(workflow.Value.SiteURL), ToCsv(workflow.Value.SiteColUrl), ToCsv(workflow.Value.DefinitionName), workflow.Value.ConsiderUpgradingToFlow, ToCsv(workflow.Value.Version), ToCsv(workflow.Value.Scope), workflow.Value.HasSubscriptions, workflow.Value.Enabled, workflow.Value.IsOOBWorkflow,
                                                                                           ToCsv(workflow.Value.ListTitle), ToCsv(workflow.Value.ListUrl), workflow.Value.ListId.ToString(), ToCsv(workflow.Value.ContentTypeName), ToCsv(workflow.Value.ContentTypeId),
                                                                                           ToCsv(workflow.Value.RestrictToType), ToCsv(workflow.Value.DefinitionDescription), workflow.Value.DefinitionId.ToString(), ToCsv(workflow.Value.SubscriptionName), workflow.Value.SubscriptionId.ToString(),
                                                                                           ToDateString(workflow.Value.LastDefinitionEdit, this.DateFormat), ToDateString(workflow.Value.LastSubscriptionEdit, this.DateFormat),
                                                                                           workflow.Value.ActionCount, ToCsv(PublishingPageScanResult.FormatList(workflow.Value.UsedActions)), ToCsv(PublishingPageScanResult.FormatList(workflow.Value.UsedTriggers)), workflow.Value.ToFLowMappingPercentage, workflow.Value.UnsupportedActionCount, ToCsv(PublishingPageScanResult.FormatList(workflow.Value.UnsupportedActionsInFlow)),
                                                                                           ToYearString(workflow.Value.LastDefinitionEdit), ToQuarterString(workflow.Value.LastDefinitionEdit), ToMonthString(workflow.Value.LastDefinitionEdit)
                                                     )));
                }

                outStream.Flush();
                this.GeneratedFileStreams.Add("ModernizationWorkflowScanResults.csv", modernizationWorkflowScanResults);
            }

            if (Options.IncludeInfoPath(this.Mode))
            {
                // Telemetry
                if (this.ScannerTelemetry != null)
                {
                    this.ScannerTelemetry.LogInfoPathScan(this.InfoPathScanResults);
                }

                MemoryStream modernizationInfoPathScanResults = new MemoryStream();
                outputHeaders = new string[] { "Site Url", "Site Collection Url", "InfoPath Usage", "Enabled", "Last user modified date", "Item count",
                                               "List Title", "List Url", "List Id", "Template",
                                               "Change Year", "Change Quarter", "Change Month"  };

                outStream = new StreamWriter(modernizationInfoPathScanResults);

                outStream.Write(string.Format("{0}\r\n", string.Join(this.Separator, outputHeaders)));
                foreach (var infoPath in this.InfoPathScanResults)
                {
                    outStream.Write(string.Format("{0}\r\n", string.Join(this.Separator, ToCsv(infoPath.Value.SiteURL), ToCsv(infoPath.Value.SiteColUrl), ToCsv(infoPath.Value.InfoPathUsage), infoPath.Value.Enabled, ToDateString(infoPath.Value.LastItemUserModifiedDate, this.DateFormat), infoPath.Value.ItemCount,
                                                                                           ToCsv(infoPath.Value.ListTitle), ToCsv(infoPath.Value.ListUrl), infoPath.Value.ListId.ToString(), ToCsv(infoPath.Value.InfoPathTemplate),
                                                                                           ToYearString(infoPath.Value.LastItemUserModifiedDate), ToQuarterString(infoPath.Value.LastItemUserModifiedDate), ToMonthString(infoPath.Value.LastItemUserModifiedDate)
                                                     )));
                }
                outStream.Flush();
                this.GeneratedFileStreams.Add("ModernizationInfoPathScanResults.csv", modernizationInfoPathScanResults);
            }

            if (Options.IncludeBlog(this.Mode))
            {
                // Telemetry
                if (this.ScannerTelemetry != null)
                {
                    this.ScannerTelemetry.LogBlogScan(this.BlogWebScanResults, this.BlogPageScanResults);
                }

                MemoryStream modernizationBlogWebScanResults = new MemoryStream();
                outputHeaders = new string[] { "Site Url", "Site Collection Url", "Web Relative Url", "Blog Type", "Web Template", "Language",
                                               "Blog Page Count", "Last blog change date", "Last blog publish date",
                                               "Change Year", "Change Quarter", "Change Month" };

                outStream = new StreamWriter(modernizationBlogWebScanResults);
                outStream.Write(string.Format("{0}\r\n", string.Join(this.Separator, outputHeaders)));
                foreach (var blogWeb in this.BlogWebScanResults)
                {
                    outStream.Write(string.Format("{0}\r\n", string.Join(this.Separator, ToCsv(blogWeb.Value.SiteURL), ToCsv(blogWeb.Value.SiteColUrl), ToCsv(blogWeb.Value.WebRelativeUrl), blogWeb.Value.BlogType, blogWeb.Value.WebTemplate, blogWeb.Value.Language,
                                                                                           blogWeb.Value.BlogPageCount, ToDateString(blogWeb.Value.LastRecentBlogPageChange, this.DateFormat), ToDateString(blogWeb.Value.LastRecentBlogPagePublish, this.DateFormat),
                                                                                           ToYearString(blogWeb.Value.LastRecentBlogPageChange), ToQuarterString(blogWeb.Value.LastRecentBlogPageChange), ToMonthString(blogWeb.Value.LastRecentBlogPageChange)
                                                     )));
                }
                outStream.Flush();
                this.GeneratedFileStreams.Add("ModernizationBlogWebScanResults.csv", modernizationBlogWebScanResults);


                MemoryStream modernizationBlogPageScanResults = new MemoryStream();
                outputHeaders = new string[] { "Site Url", "Site Collection Url", "Web Relative Url", "Blog Type", "Page Relative Url", "Page Title",
                                               "Modified At", "Modified By", "Published At" };

                outStream = new StreamWriter(modernizationBlogPageScanResults);
                outStream.Write(string.Format("{0}\r\n", string.Join(this.Separator, outputHeaders)));
                foreach (var blogPage in this.BlogPageScanResults)
                {
                    outStream.Write(string.Format("{0}\r\n", string.Join(this.Separator, ToCsv(blogPage.Value.SiteURL), ToCsv(blogPage.Value.SiteColUrl), ToCsv(blogPage.Value.WebRelativeUrl), blogPage.Value.BlogType ,ToCsv(blogPage.Value.PageRelativeUrl), ToCsv(blogPage.Value.PageTitle),
                                                                                           ToDateString(blogPage.Value.ModifiedAt, this.DateFormat), ToCsv(blogPage.Value.ModifiedBy), ToDateString(blogPage.Value.PublishedDate, this.DateFormat)
                                                     )));
                }
                outStream.Flush();
                this.GeneratedFileStreams.Add("ModernizationBlogPageScanResults.csv", modernizationBlogPageScanResults);
            }

            VersionWarning();

            Console.WriteLine("=====================================================");
            Console.WriteLine("All done. Took {0} for {1} sites", (DateTime.Now - start).ToString(), this.ScannedSites);
            Console.WriteLine("=====================================================");

            return start;
        }

        private void VersionWarning()
        {
            if (!string.IsNullOrEmpty(this.NewVersion))
            {
                var currentColor = Console.ForegroundColor;
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"Scanner version {this.NewVersion} is available. You're currently running {this.CurrentVersion}.");
                Console.WriteLine($"Download the latest version of the scanner from {VersionCheck.newVersionDownloadUrl}");
                Console.ForegroundColor = currentColor;
            }
        }
        #endregion

    }
}
