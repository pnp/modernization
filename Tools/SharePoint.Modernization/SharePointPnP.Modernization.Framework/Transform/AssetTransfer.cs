using SharePointPnP.Modernization.Framework.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointPnP.Modernization.Framework.Transform
{
    /// <summary>
    /// Class for operations for transferring the assets over to the target site collection
    /// </summary>
    public class AssetTransfer
    {
        //Plan:
        //  Detect for referenced assets within the web parts
        //  Referenced assets should only be files e.g. not aspx pages and located in the pages, site pages libraries
        //  Ensure the referenced assets exist within the same site collection/web according to the level of transformation
        //  With the modern destination, locate assets in the site assets library with in a folder using the same naming convention as SharePoint Comm Sites
        //  Add copy assets method to transfer the files to target site collection
        //  Store a dictionary of copied assets to update the URLs of the transferred web parts
        //  Phased approach for this: 
        //      Image Web Parts
        //      Text Web Parts with inline images (need to determine how they are handled)
        //      TBC - expanded as testing progresses

        /// <summary>
        /// Collect urls to referenced resources
        /// </summary>
        /// <returns></returns>
        public List<AssetTransferReferenceEntity> CollectUrlReferencesFromWebPartContent()
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Checks if the URL is located in a supported location
        /// </summary>
        public bool IsAssetInSupportedLocation(string currentContextUrl, string sourceUrl)
        {
            //  Referenced assets should only be files e.g. 
            //      not aspx pages 
            //      located in the pages, site pages libraries
            //  Ensure the referenced assets exist within the same site collection/web according to the level of transformation

            throw new NotImplementedException();
        }

        /// <summary>
        /// Ensure the site assets and page sub-folder exists in the target location
        /// </summary>
        public void EnsureDestination()
        {
            // We cannot assume the SiteAssets library exists, in the case of vanilla communication sites - provision a new library if none exists
            // If a site assets library exist, add a folder, into the library using the same format as SharePoint uses for creating sub folders for pages

            throw new NotImplementedException();
        }

        /// <summary>
        /// Create a site assets library
        /// </summary>
        public void CreateSiteAssetsLibrary()
        {
            // Use a PnP Provisioning template to create a site assets library

            throw new NotImplementedException();
        }

        /// <summary>
        /// Copy the file from the source to the target location
        /// </summary>
        public void CopyAssetToTargetLocation()
        {
            // This copies the latest version of the asset to the target site collection
            throw new NotImplementedException();
        }

        /// <summary>
        /// Stores an asset transfer reference
        /// </summary>
        /// <param name="assetTransferReferenceEntity"></param>
        /// <param name="update"></param>
        public void StoreAssetTransferReference(AssetTransferReferenceEntity assetTransferReferenceEntity, bool? update)
        {
            // Using the Cache Manager store the asset transfer references
            // If update - treat the source URL as unique, if multiple web parts reference to this, then it will still refer to the single resource
            throw new NotImplementedException();
        }

        /// <summary>
        /// Get all asset transfer references
        /// </summary>
        public void GetAssetTransferReferences()
        {
            // Using the Cache Manager retrieve asset transfer references (all)
            throw new NotImplementedException();
        }

        /// <summary>
        /// Gets a list of assets pending transfer to the target location
        /// </summary>
        /// <returns></returns>
        public List<AssetTransferReferenceEntity> GetPendingAssetTransfers()
        {
            // Using the Cache Manager get the assets new transferred
            throw new NotImplementedException();
        }

    }
}
