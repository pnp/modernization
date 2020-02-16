using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointPnP.Modernization.Framework.Entities
{
    [Serializable]
    public class TermData
    {
        /// <summary>
        /// Term Guid
        /// </summary>
        public Guid TermGuid { get; set; }

        /// <summary>
        /// Term Label
        /// </summary>
        public string TermLabel { get; set; }

        /// <summary>
        /// Term Path
        /// </summary>
        public string TermPath { get; set; }

        /// <summary>
        /// Term Set ID
        /// </summary>
        public Guid TermSetId { get; set; }
    }
}
