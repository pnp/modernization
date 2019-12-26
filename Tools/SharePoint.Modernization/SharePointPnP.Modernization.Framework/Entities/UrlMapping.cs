using System;

namespace SharePointPnP.Modernization.Framework.Entities
{
    /// <summary>
    /// Url map entry
    /// </summary>
    [Serializable]
    public class UrlMapping
    {
        /// <summary>
        /// Url to be replaced
        /// </summary>
        public string SourceUrl { get; set; }

        /// <summary>
        /// Url replacement value
        /// </summary>
        public string TargetUrl { get; set; }
    }
}
