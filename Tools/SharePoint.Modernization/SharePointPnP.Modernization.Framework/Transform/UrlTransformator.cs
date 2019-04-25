using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Utilities;
using SharePointPnP.Modernization.Framework.Cache;
using SharePointPnP.Modernization.Framework.Telemetry;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace SharePointPnP.Modernization.Framework.Transform
{
    public class UrlTransformator : BaseTransform
    {
        private ClientContext sourceContext;
        private ClientContext targetContext;

        private string sourceSiteUrl;
        private string sourceWebUrl;
        private string targetWebUrl;
        private string pagesLibrary;

        #region Construction
        public UrlTransformator(ClientContext sourceContext, ClientContext targetContext, IList<ILogObserver> logObservers = null)
        {
            // Hookup logging
            if (logObservers != null)
            {
                foreach (var observer in logObservers)
                {
                    base.RegisterObserver(observer);
                }
            }

            // Ensure source and target context are set
            if (sourceContext == null && targetContext != null)
            {
                sourceContext = targetContext;
            }

            if (targetContext == null && sourceContext != null)
            {
                targetContext = sourceContext;
            }

            // Grab the needed information to drive url rewrite
            this.sourceContext = sourceContext;
            this.targetContext = targetContext;
            this.sourceSiteUrl = sourceContext.Site.EnsureProperty(p => p.Url);
            this.sourceWebUrl = sourceContext.Web.EnsureProperty(p => p.Url);
            this.pagesLibrary = CacheManager.Instance.GetPublishingPagesLibraryName(this.sourceContext);
            this.targetWebUrl = targetContext.Web.EnsureProperty(p => p.Url);
        }
        #endregion

        public string Transform(string input)
        {
            // Do we need to rewrite?
            if (this.sourceWebUrl.Equals(this.targetWebUrl, StringComparison.InvariantCultureIgnoreCase))
            {
                return input;
            }

            return ReWriteUrls(input, this.sourceSiteUrl, this.sourceWebUrl, this.targetWebUrl, this.pagesLibrary);
        }

        private string ReWriteUrls(string input, string sourceSiteUrl, string sourceWebUrl, string targetWebUrl, string pagesLibrary)
        {
            //TODO: find a solution for managed navigation links as they're returned as "https://bertonline.sharepoint.com/sites/ModernizationTarget/_layouts/15/FIXUPREDIRECT.ASPX?WebId=b710de6c-ff13-41f2-b119-0e7ad57269d2&TermSetId=c6eba345-eaf4-4e17-9c3e-c8436e017326&TermId=c2d20b8f-e70b-417d-8aa3-d5e3b59f6167"

            bool isSubSite = !sourceSiteUrl.Equals(sourceWebUrl, StringComparison.InvariantCultureIgnoreCase);

            // Rewrite url's from pages library to sitepages
            if (!string.IsNullOrEmpty(pagesLibrary))
            {
                string pagesSourceWebUrl = UrlUtility.Combine(sourceWebUrl, pagesLibrary);
                string sitePagesTargetWebUrl = UrlUtility.Combine(targetWebUrl, "sitepages");

                if (pagesSourceWebUrl.StartsWith("https://", StringComparison.InvariantCultureIgnoreCase))
                {
                    input = RewriteUrl(input, pagesSourceWebUrl, sitePagesTargetWebUrl);

                    // Make relative for next replacement attempt
                    pagesSourceWebUrl = MakeRelative(pagesSourceWebUrl);
                    sitePagesTargetWebUrl = MakeRelative(sitePagesTargetWebUrl);
                }

                input = RewriteUrl(input, pagesSourceWebUrl, sitePagesTargetWebUrl);
            }

            // Rewrite web urls
            if (sourceWebUrl.StartsWith("https://", StringComparison.InvariantCultureIgnoreCase))
            {
                input = RewriteUrl(input, sourceWebUrl, targetWebUrl);

                // Make relative for next replacement attempt
                sourceWebUrl = MakeRelative(sourceWebUrl);
                targetWebUrl = MakeRelative(targetWebUrl);
            }

            input = RewriteUrl(input, sourceWebUrl, targetWebUrl);

            if (isSubSite)
            {
                // Rewrite url's from pages library to sitepages
                if (!string.IsNullOrEmpty(pagesLibrary))
                {
                    string pagesSourceSiteUrl = UrlUtility.Combine(sourceSiteUrl, pagesLibrary);
                    string sitePagesTargetWebUrl = UrlUtility.Combine(targetWebUrl, "sitepages");

                    if (pagesSourceSiteUrl.StartsWith("https://", StringComparison.InvariantCultureIgnoreCase))
                    {
                        input = RewriteUrl(input, pagesSourceSiteUrl, sitePagesTargetWebUrl);

                        // Make relative for next replacement attempt
                        pagesSourceSiteUrl = MakeRelative(pagesSourceSiteUrl);
                        sitePagesTargetWebUrl = MakeRelative(sitePagesTargetWebUrl);
                    }

                    input = RewriteUrl(input, pagesSourceSiteUrl, sitePagesTargetWebUrl);
                }

                // Rewrite root site urls
                if (sourceSiteUrl.StartsWith("https://", StringComparison.InvariantCultureIgnoreCase))
                {
                    input = RewriteUrl(input, sourceSiteUrl, targetWebUrl);

                    // Make relative for next replacement attempt
                    sourceSiteUrl = MakeRelative(sourceSiteUrl);
                    targetWebUrl = MakeRelative(targetWebUrl);
                }

                input = RewriteUrl(input, sourceSiteUrl, targetWebUrl);
            }

            return input;
        }

        private string RewriteUrl(string input, string from, string to)
        {
            var regex = new Regex($"{from}", RegexOptions.IgnoreCase);
            if (regex.IsMatch(input))
            {
                input = regex.Replace(input, to);
            }

            return input;
        }

        private string MakeRelative(string url)
        {
            Uri uri = new Uri(url);
            return uri.AbsolutePath;
        }

    }
}
