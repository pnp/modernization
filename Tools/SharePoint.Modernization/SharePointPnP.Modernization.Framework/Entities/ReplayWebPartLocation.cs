using OfficeDevPnP.Core.Pages;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointPnP.Modernization.Framework.Entities
{
    /// <summary>
    /// Entity Class for storing web part location data
    /// </summary>
    [Serializable]
    public class ReplayWebPartLocation
    {

        public ReplayWebPartLocation()
        {
            Row = int.MinValue;
            Column = int.MinValue;
            Order = int.MinValue;
            MovedToColumn = int.MinValue;
            MovedToOrder = int.MinValue;
            MovedToRow = int.MinValue;
        }

        public string PageUrl { get; set; }
        
        public Guid SourceWebPartId { get; set; }
        
        public string SourceWebPartType { get; set; }
        
        public string TargetWebPartTypeId { get; set; }
        
        public Guid TargetWebPartInstanceId { get; set; }

        public int Row { get; set; }
        
        public int Column { get; set; }
        
        public int Order { get; set; }

        public int MovedToRow { get; set; }

        public int MovedToColumn { get; set; }

        public int MovedToOrder { get; set; }

    }
}
