using SharePointPnP.Modernization.Framework.Telemetry;
using System;
using System.Collections.Generic;

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
            TargetPage = string.Empty;
            SourceSite = string.Empty;
            TargetSite = string.Empty;
            BaseSourceUrl = string.Empty;
            BaseTargetUrl = string.Empty;
            AssetsTransferred = new List<Tuple<LogLevel, LogEntry>>();
            PageLogsOrdered = new List<Tuple<LogLevel, LogEntry>>();
            TransformationVerboseSummary = new List<Tuple<LogLevel, LogEntry>>();
            TransformationVerboseDetails = new List<Tuple<LogLevel, LogEntry>>();
            TransformationSettings = new List<Tuple<string, string>>();
        }

        /// <summary>
        /// Source Page
        /// </summary>
        public string SourcePage { get; set; }

        public string SourceSite { get; set; }


        public string TargetPage { get; set; }

        public string TargetSite { get; set; }

        /// <summary>
        /// Date report generated
        /// </summary>
        public DateTime ReportDate { get; set; }

        public string BaseSourceUrl { get; set; }
        public string BaseTargetUrl { get; set; }

        public TimeSpan TransformationDuration { get; set; }

        public bool IsFirstAnalysis { get; set; }


        public string PageId { get; set; }


        public List<Tuple<LogLevel, LogEntry>> AssetsTransferred { get; set; }



        /// <summary>
        /// List of warnings raised
        /// </summary>
        public List<Tuple<LogLevel, LogEntry>> Warnings { get; set; }

        /// <summary>
        /// List of errors raised
        /// </summary>
        public List<Tuple<LogLevel, LogEntry>> Errors { get; set; }

        /// <summary>
        /// List of critical application error
        /// </summary>
        public List<Tuple<LogLevel, LogEntry>> CriticalErrors { get; set; }



        /// <summary>
        /// Page Logs ordered
        /// </summary>
        public List<Tuple<LogLevel, LogEntry>> PageLogsOrdered { get; set; }

        /// <summary>
        /// Logs that contain summary data for verbose logging
        /// </summary>
        public List<Tuple<LogLevel, LogEntry>> TransformationVerboseSummary { get; set; }

        public List<Tuple<LogLevel, LogEntry>> TransformationVerboseDetails { get; set; }

        public List<Tuple<string, string>> TransformationSettings { get; set; }
    }
}