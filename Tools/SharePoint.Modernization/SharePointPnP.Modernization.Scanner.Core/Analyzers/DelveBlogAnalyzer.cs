using Microsoft.SharePoint.Client;
using SharePoint.Modernization.Scanner.Core;
using SharePoint.Modernization.Scanner.Core.Analyzers;
using SharePoint.Modernization.Scanner.Core.Results;
using SharePointPnP.Modernization.Framework.Cache;
using System;
using System.Collections.Generic;

namespace SharePointPnP.Modernization.Scanner.Core.Analyzers
{
    public class DelveBlogAnalyzer: BaseAnalyzer
    {

        public readonly string keyDelveSitesList = "DelveSitesList";

        #region Construction
        /// <summary>
        /// Delve Blog analyzer construction
        /// </summary>
        /// <param name="url">Url of the web to be analyzed</param>
        /// <param name="siteColUrl">Url of the site collection hosting this web</param>
        /// <param name="scanJob">Job that launched this analyzer</param>
        public DelveBlogAnalyzer(string url, string siteColUrl, ModernizationScanJob scanJob) : base(url, siteColUrl, scanJob)
        {
        }
        #endregion

        #region Analysis
        public override TimeSpan Analyze(ClientContext cc)
        {

            // Get site usage information
            List<string> propertiesToRetrieve = new List<string>
                    {
                        "LastModifiedTime",
                        "ModifiedBy",
                        "DetectedLanguage",
                        "SPWebUrl",
                        "Path",
                        "Title",
                        "ViewsRecent",
                        "ViewsRecentUniqueUsers",
                        "ViewsLifeTime",
                        "ViewsLifeTimeUniqueUsers"
                    };

            Uri rootSite = new Uri(this.SiteCollectionUrl);
            string pathFilter = $"{rootSite.Scheme}://{rootSite.DnsSafeHost}/portals/personal/*";

            Dictionary<string, BlogWebScanResult> tempWebResults = new Dictionary<string, BlogWebScanResult>();

            // try get the delve search results from cache as they're always the same when used in a distributed scan scenario
            List<Dictionary<string, string>> results;
            var delveSitesList = this.ScanJob.Store.Get<List<Dictionary<string, string>>>(this.ScanJob.StoreOptions.GetKey(keyDelveSitesList));
            if (delveSitesList != null)
            {
                results = new List<Dictionary<string, string>>();
                results.AddRange(delveSitesList);
            }
            else
            {
                results = this.ScanJob.Search(cc.Web, $"path:{pathFilter} AND ContentTypeId:\"0x010100DA3A7E6E3DB34DFF8FDEDE1F4EBAF95D*\"", propertiesToRetrieve);
                if (results != null)
                {
                    this.ScanJob.Store.Set<List<Dictionary<string, string>>>(this.ScanJob.StoreOptions.GetKey(keyDelveSitesList), results, this.ScanJob.StoreOptions.EntryOptions);
                }
            }

            if (results != null && results.Count > 0)
            {

                foreach(var result in results)
                {
                    string url = result["SPWebUrl"]?.ToLower().Replace($"{rootSite.Scheme}://{rootSite.DnsSafeHost}".ToLower(), "");

                    DateTime lastModified = DateTime.MinValue;

                    if (result["LastModifiedTime"] != null)
                    {
                        DateTime.TryParse(result["LastModifiedTime"], out lastModified);
                    }

                    BlogWebScanResult scanResult = null;
                    BlogPageScanResult blogPageScanResult;
                    if (tempWebResults.ContainsKey(url))
                    {
                        // Increase page counter
                        tempWebResults[url].BlogPageCount += 1;

                        // Build page record
                        blogPageScanResult = AddBlogPageResult($"{rootSite.Scheme}://{rootSite.DnsSafeHost}".ToLower(), url, result, lastModified);
                    }
                    else
                    {
                        scanResult = new BlogWebScanResult
                        {
                            SiteColUrl = result["SPWebUrl"],
                            SiteURL = result["SPWebUrl"],
                            WebRelativeUrl = url,
                            WebTemplate = "POINTPUBLISHINGPERSONAL#0",
                            BlogType = BlogType.Delve,
                            BlogPageCount = 1,
                            LastRecentBlogPageChange = lastModified,
                            LastRecentBlogPagePublish = lastModified,
                            Language = 1033
                        };

                        tempWebResults.Add(url, scanResult);

                        // Build page record
                        blogPageScanResult = AddBlogPageResult($"{rootSite.Scheme}://{rootSite.DnsSafeHost}".ToLower(), url, result, lastModified);
                    }

                    if (blogPageScanResult != null)
                    {
                        if (!this.ScanJob.BlogPageScanResults.TryAdd($"blogScanResult.PageURL.{Guid.NewGuid()}", blogPageScanResult))
                        {
                            ScanError error = new ScanError()
                            {
                                Error = $"Could not add delve blog page scan result for {blogPageScanResult.SiteColUrl}",
                                SiteColUrl = this.SiteCollectionUrl,
                                SiteURL = this.SiteUrl,
                                Field1 = "DelveBlogPageAnalyzer",
                            };
                            this.ScanJob.ScanErrors.Push(error);
                        }
                    }
                }

                // Copy the temp scan results to the actual structure
                foreach(var blogWebResult in tempWebResults)
                {
                    if (!this.ScanJob.BlogWebScanResults.TryAdd($"blogScanResult.WebURL.{Guid.NewGuid()}", blogWebResult.Value))
                    {
                        ScanError error = new ScanError()
                        {
                            Error = $"Could not add delve blog web scan result for {blogWebResult.Value.SiteColUrl}",
                            SiteColUrl = this.SiteCollectionUrl,
                            SiteURL = this.SiteUrl,
                            Field1 = "DelveBlogWebAnalyzer",
                        };
                        this.ScanJob.ScanErrors.Push(error);
                    }
                }

            }

            return base.Analyze(cc);
        }
        #endregion

        #region Helper methods
        private BlogPageScanResult AddBlogPageResult(string rootUrl, string url, Dictionary<string, string> result, DateTime lastModified)
        {
            string pageTitle = "";
            if (result["Title"] != null)
            {
                pageTitle = result["Title"];
            }

            string pageRelativeUrl = "";
            if  (result["Path"] != null)
            {
                pageRelativeUrl = result["Path"].ToLower().Replace(rootUrl, "");
            }

            string modifiedBy = "";
            if (result["ModifiedBy"] != null)
            {
                modifiedBy = result["ModifiedBy"];
            }

            BlogPageScanResult scanResult = new BlogPageScanResult
            {
                SiteColUrl = result["SPWebUrl"],
                SiteURL = result["SPWebUrl"],
                WebRelativeUrl = url,
                PageRelativeUrl = pageRelativeUrl,
                BlogType = BlogType.Delve,
                PageTitle = pageTitle,
                ModifiedAt = lastModified,
                PublishedDate = lastModified,
                ModifiedBy = modifiedBy,
            };

            return scanResult;
        }

        #endregion
    }
}
