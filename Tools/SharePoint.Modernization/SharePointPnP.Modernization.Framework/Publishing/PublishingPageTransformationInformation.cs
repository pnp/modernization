using Microsoft.SharePoint.Client;
using SharePointPnP.Modernization.Framework.Transform;
using System.Collections.Generic;

namespace SharePointPnP.Modernization.Framework.Publishing
{
    /// <summary>
    /// Information used to configure the publishing page transformation process
    /// </summary>
    public class PublishingPageTransformationInformation: BaseTransformationInformation
    {
        #region Construction
        /// <summary>
        /// Instantiates the page transformation class
        /// </summary>
        /// <param name="sourcePage">Page we want to transform</param>
        public PublishingPageTransformationInformation(ListItem sourcePage) : this(sourcePage, false)
        {
        }

        /// <summary>
        /// Instantiates the page transformation class
        /// </summary>
        /// <param name="sourcePage">Page we want to transform</param>
        /// <param name="overwrite">Do we overwrite the target page if it already exists</param>
        public PublishingPageTransformationInformation(ListItem sourcePage, bool overwrite)
        {
            SourcePage = sourcePage;
            Overwrite = overwrite;
            HandleWikiImagesAndVideos = true;
            KeepPageSpecificPermissions = true;
            CopyPageMetadata = false;
            SkipTelemetry = false;
            RemoveEmptySectionsAndColumns = true;
            // Populate with OOB mapping properties
            MappingProperties = new Dictionary<string, string>(5)
            {
                { "UseCommunityScriptEditor", "false" },
                { "SummaryLinksToQuickLinks", "true" }
            };
        }
        #endregion

        #region Page Properties
        #endregion
    }
}
