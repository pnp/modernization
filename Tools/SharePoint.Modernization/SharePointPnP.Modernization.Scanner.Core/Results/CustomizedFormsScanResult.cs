using Microsoft.SharePoint.Client;
using System;

namespace SharePoint.Modernization.Scanner.Core.Results
{
    public class CustomizedFormsScanResult : Scan
    {
        /// <summary>
        /// Type of form page
        /// </summary>
        public PageType FormType { get; set; }

        /// <summary>
        /// Id of the form page
        /// </summary>
        public Guid PageId { get; set; }

        /// <summary>
        /// Url of the form page
        /// </summary>
        public string Url { get; set; }

        /// <summary>
        /// Id of the web part on the form page
        /// </summary>
        public Guid WebpartId { get; set; }

    }
}
