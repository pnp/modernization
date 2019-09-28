using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointPnP.Modernization.Framework.Entities
{
    public class UserMappingEntity
    {
        /// <summary>
        /// Source user reference
        /// </summary>
        public string SourceUser { get; set; }

        /// <summary>
        /// Target user reference
        /// </summary>
        public string TargetUser { get; set; }
    }
}
