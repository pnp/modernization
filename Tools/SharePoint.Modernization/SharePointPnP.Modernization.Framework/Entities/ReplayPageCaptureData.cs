using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointPnP.Modernization.Framework.Entities
{
    [Serializable]
    public class ReplayPageCaptureData
    {
        /// <summary>
        /// Constructor for replay page capture data
        /// </summary>
        public ReplayPageCaptureData()
        {
            ReplayWebPartLocations = new List<ReplayWebPartLocation>();
        }

        public string PageName { get; set; }

        public string PageLayoutName { get; set; }

        public string PageUrl { get; set; }

        public List<ReplayWebPartLocation> ReplayWebPartLocations { get;set; }
    }
}
