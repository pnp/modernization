using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointPnP.Modernization.Framework.Entities
{
    /// <summary>
    /// Model for asset transfer status for references
    /// </summary>
    public class AssetTransferredEntity
    {
        /// <summary>
        /// Source web part URL reference - for checking if transfer occurred before
        /// </summary>
        public string SourceAssetUrl { get; set; }

        /// <summary>
        /// Target web part URL reference - for checking if transfer occurred before
        /// </summary>
        public string TargetAssetFolderUrl { get; set; }

        /// <summary>
        /// Store the final URL for the asset that has been transferred
        /// </summary>
        public string TargetAssetTransferredUrl { get; set; }

    }
}
