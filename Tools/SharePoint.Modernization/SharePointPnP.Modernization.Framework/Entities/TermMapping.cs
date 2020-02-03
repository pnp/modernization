using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointPnP.Modernization.Framework.Entities
{
    public class TermMapping
    {
        public string SourceRaw { get; set; }
        public string TargetRaw { get; set; }

        public string SourceTermPath { get; set; }
        public string TargetTermPath { get; set; }

        public Guid SourceTermId { get; set; }

        public Guid TargetTermId { get; set; }
    }
}
