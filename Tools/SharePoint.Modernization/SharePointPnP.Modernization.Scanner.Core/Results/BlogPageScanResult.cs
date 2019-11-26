using System;

namespace SharePoint.Modernization.Scanner.Core.Results
{
    /// <summary>
    /// Stores information about a found blog page
    /// </summary>
    public class BlogPageScanResult: Scan
    {

        public BlogPageScanResult()
        {
            this.BlogType = BlogType.Classic;
        }

        /// <summary>
        /// Type of blog page
        /// </summary>
        public BlogType BlogType { get; set; }
        
        /// <summary>
        /// Web relative Url
        /// </summary>
        public string WebRelativeUrl { get; set; }

        /// <summary>
        /// page relative Url
        /// </summary>
        public string PageRelativeUrl { get; set; }

        /// <summary>
        /// Title of the scanned blog page
        /// </summary>
        public string PageTitle { get; set; }

        // Page modification information
        public DateTime ModifiedAt { get; set; }
        public string ModifiedBy { get; set; }

        /// <summary>
        /// Blog publishing date
        /// </summary>
        public DateTime PublishedDate { get; set; }

    }
}
