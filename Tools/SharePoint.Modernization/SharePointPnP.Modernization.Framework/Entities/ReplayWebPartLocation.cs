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
            SourceWebPartId = Guid.Empty;
            TargetWebPartInstanceId = Guid.Empty;
            Row = int.MinValue;
            Column = int.MinValue;
            Order = int.MinValue;
            MovedToColumn = int.MinValue;
            MovedToOrder = int.MinValue;
            MovedToRow = int.MinValue;
        }

        public Guid SourceWebPartId { get; set; }
        
        public string SourceWebPartType { get; set; }

        public string SourceWebPartTitle { get; set; }
        
        public string TargetWebPartTypeId { get; set; }
        
        public Guid TargetWebPartInstanceId { get; set; }

        public int Row { get; set; }
        
        public int Column { get; set; }
        
        public int ColumnFactor { get; set; }

        public int Order { get; set; }

        public int MovedToRow { get; set; }

        public int MovedToColumn { get; set; }

        public int MovedToOrder { get; set; }

        public int MovedToColumnFactor { get; set; }

        public bool MovedToIsVerticalColumn { get; set; }

        public int MovedToRowZoneEmphesis { get; set; }

        /// <summary>
        /// Can use move to location information
        /// </summary>
        public bool CanUseMoveToLocation
        {
            get
            {
                if(MovedToRow != int.MinValue && MovedToColumn != int.MinValue && MovedToOrder != int.MinValue)
                {
                    return true;
                }

                return false;
            }
        }

        public string SourceGroupName { get; set; }

    }
}
