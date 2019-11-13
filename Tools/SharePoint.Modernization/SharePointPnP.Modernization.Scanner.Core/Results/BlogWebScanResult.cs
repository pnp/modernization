using System;

namespace SharePoint.Modernization.Scanner.Core.Results
{
    public class BlogWebScanResult: Scan
    {
        public BlogWebScanResult()
        {
            this.LastRecentBlogPageChange = DateTime.MinValue;
            this.LastRecentBlogPagePublish = DateTime.MinValue;
        }

        /// <summary>
        /// Web relative Url
        /// </summary>
        public string WebRelativeUrl { get; set; }

        /// <summary>
        /// Web template (e.g. STS#0)
        /// </summary>
        public string WebTemplate { get; set; }

        /// <summary>
        /// Language of the used blog site
        /// </summary>
        public uint Language { get; set; }

        /// <summary>
        /// Number of blog pages in this site
        /// </summary>
        public int BlogPageCount { get; set; }

        /// <summary>
        /// Most recent blog change date
        /// </summary>
        public DateTime LastRecentBlogPageChange { get; set; }

        /// <summary>
        /// Most recent blog publish date
        /// </summary>
        public DateTime LastRecentBlogPagePublish { get; set; }

    }
}
