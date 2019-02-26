using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointPnP.Modernization.Framework.Telemetry
{
    public class LogEntry
    {
        /// <summary>
        /// Gets or sets Log message
        /// </summary>
        public string Message { get; set; }
        /// <summary>
        /// Gets or sets CorrelationId of type Guid
        /// </summary>
        public Guid CorrelationId { get; set; }
        /// <summary>
        /// Gets or sets Log source
        /// </summary>
        public string Source { get; set; }
        /// <summary>
        /// Gets or sets Log Exception
        /// </summary>
        public Exception Exception { get; set; }
    }
}
