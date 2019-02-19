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
    public class AssetTransferReferenceEntity
    {
        /// <summary>
        /// Value to mark the asset exists on the destination
        /// </summary>
        public bool IsAssetTransferred { get; set; }
        
        /// <summary>
        /// Source web part URL reference
        /// </summary>
        public string SourceAssetReference { get; set; }

        /// <summary>
        /// Target web part URL reference
        /// </summary>
        public string TargetAssetReference { get; set; }

        /// <summary>
        /// Is the asset in a unsupported asset location
        /// </summary>
        public bool? IsAssetInUnSupportedLocation { get; set; }
        

    }
}
