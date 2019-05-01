using SharePointPnP.Modernization.Framework.Telemetry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointPnP.Modernization.Framework.Entities
{
    public class TransformationLogAnalysis
    {
        /// <summary>
        /// Constructor for transformation report
        /// </summary>
        public TransformationLogAnalysis()
        {
            Warnings = new List<Tuple<LogLevel, LogEntry>>();
            Errors = new List<Tuple<LogLevel, LogEntry>>();
            SourcePage = string.Empty;
        }

        /// <summary>
        /// Source Page
        /// </summary>
        public string SourcePage { get; set; }

        /// <summary>
        /// List of warnings raised
        /// </summary>
        public IEnumerable<Tuple<LogLevel, LogEntry>> Warnings { get; set; }

        /// <summary>
        /// List of errors raised
        /// </summary>
        public IEnumerable<Tuple<LogLevel, LogEntry>> Errors { get; set; }

        /// <summary>
        /// List of critical application error
        /// </summary>
        public IEnumerable<Tuple<LogLevel, LogEntry>> CriticalErrors { get; set; }

    }
}
