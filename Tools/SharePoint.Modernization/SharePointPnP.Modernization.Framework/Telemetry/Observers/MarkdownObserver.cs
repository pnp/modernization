﻿using SharePointPnP.Modernization.Framework.Extensions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace SharePointPnP.Modernization.Framework.Telemetry.Observers
{
    /// <summary>
    /// Markdown observer intended for end-user output
    /// </summary>
    public class MarkdownObserver : ILogObserver
    {

        // Cache the logs between calls
        private static readonly Lazy<List<Tuple<LogLevel, LogEntry>>> _lazyLogInstance = new Lazy<List<Tuple<LogLevel, LogEntry>>>(() => new List<Tuple<LogLevel, LogEntry>>());
        protected bool _includeDebugEntries;
        protected bool _includeVerbose;
        protected DateTime _reportDate;
        protected string _reportFileName = "";
        protected string _reportFolder = Environment.CurrentDirectory;
        protected string _pageBeingTransformed;

        #region Construction
        /// <summary>
        /// Constructor for specifying to include debug entries
        /// </summary>
        /// <param name="fileName">Name used to construct the log file name</param>
        /// <param name="folder">Folder that will hold the log file</param>
        /// <param name="includeDebugEntries">Include Debug Log Entries</param>
        public MarkdownObserver(string fileName = "", string folder = "", bool includeDebugEntries = false, bool includeVerbose = false)
        {
            _includeDebugEntries = includeDebugEntries;
            _includeVerbose = includeVerbose;
            _reportDate = DateTime.Now;

            if (!string.IsNullOrEmpty(folder))
            {
                _reportFolder = folder;
            }

            // Drop possible file extension as we want to ensure we have a .md extension
            _reportFileName = System.IO.Path.GetFileNameWithoutExtension(fileName);

#if DEBUG && MEASURE && MEASURE
           _includeDebugEntries = true; //Override for debugging locally
#endif
        }
        #endregion

        #region Markdown Tokens
        private const string Heading1 = "#";
        private const string Heading2 = "##";
        private const string Heading3 = "###";
        private const string Heading4 = "####";
        private const string Heading5 = "#####";
        private const string Heading6 = "######";
        private const string UnorderedListItem = "-";
        private const string Italic = "_";
        private const string Bold = "**";
        private const string BlockQuotes = "> ";
        private const string TableHeaderColumn = "-------------";
        private const string TableColumnSeperator = " | ";
        private const string Link = "[{0}]({1})";
        #endregion

        /// <summary>
        /// Get the single List<LogEntry> instance, singleton pattern
        /// </summary>
        public static List<Tuple<LogLevel, LogEntry>> Logs
        {
            get
            {
                return _lazyLogInstance.Value;
            }
        }

        /// <summary>
        /// Debug level of data not recorded unless in debug mode
        /// </summary>
        /// <param name="entry"></param>
        public void Debug(LogEntry entry)
        {
            if (_includeDebugEntries)
            {
                entry.PageId = this._pageBeingTransformed;
                Logs.Add(new Tuple<LogLevel, LogEntry>(LogLevel.Debug, entry));
            }
        }

        /// <summary>
        /// Errors 
        /// </summary>
        /// <param name="entry"></param>
        public void Error(LogEntry entry)
        {
            entry.PageId = this._pageBeingTransformed;
            Logs.Add(new Tuple<LogLevel, LogEntry>(LogLevel.Error, entry));
        }

        /// <summary>
        /// Reporting operations throughout the transform process
        /// </summary>
        /// <param name="entry"></param>
        public void Info(LogEntry entry)
        {
            entry.PageId = this._pageBeingTransformed;
            Logs.Add(new Tuple<LogLevel, LogEntry>(LogLevel.Information, entry));
        }

        /// <summary>
        /// Report on any warnings generated by the reporting tool
        /// </summary>
        /// <param name="entry"></param>
        public void Warning(LogEntry entry)
        {
            entry.PageId = this._pageBeingTransformed;
            Logs.Add(new Tuple<LogLevel, LogEntry>(LogLevel.Warning, entry));
        }

        /// <summary>
        /// Sets the id of the page that's being transformed
        /// </summary>
        /// <param name="pageId">Id of the page</param>
        public void SetPageId(string pageId)
        {
            this._pageBeingTransformed = pageId;
        }

        /// <summary>
        /// Generates a markdown based report based on the logs
        /// </summary>
        /// <returns></returns>
        protected virtual string GenerateReport(bool includeHeading = true)
        {
            StringBuilder report = new StringBuilder();

            // Get one log entry per page...assumes that this log entry is included by each transformator
            var distinctLogs = Logs.Where(p => p.Item2.Heading == LogStrings.Heading_Summary && p.Item2.Significance == LogEntrySignificance.SourceSiteUrl); //TODO: Need to improve this

            bool first = true;
            foreach (var distinctLogEntry in distinctLogs)
            {
                var logEntriesToProcess = Logs.Where(p => p.Item2.PageId == distinctLogEntry.Item2.PageId);
                GenerateReportForPage(report, logEntriesToProcess, first, _includeVerbose);

                first = false;
            }

            //TODO: For Summary Mode only
            if (!_includeVerbose)
            {

                //if errors - include just errors

                //if warnings - include just warnings

                //if critical - include critical messages
            }

            return report.ToString();
        }

        /// <summary>
        /// Generates a markdown based report based on the logs
        /// </summary>
        /// <returns></returns>
        private string GenerateReportForPage(StringBuilder report, IEnumerable<Tuple<LogLevel, LogEntry>> logEntriesToProcess, bool firstRun = true, bool includeVerbose = false)
        {

            // Log Analysis

            // This could display something cool here e.g. Time taken to transform and transformation options e.g. PageTransformationInformation details
            var reportDate = _reportDate;
            var allLogs = logEntriesToProcess.OrderBy(l => l.Item2.EntryTime);
            var transformationSummary = allLogs.Where(l => l.Item2.Heading == LogStrings.Heading_Summary);

            var sourcePage = transformationSummary.FirstOrDefault(l => l.Item2.Significance == LogEntrySignificance.SourcePage);
            var targetPage = transformationSummary.FirstOrDefault(l => l.Item2.Significance == LogEntrySignificance.TargetPage);
            var sourceSite = transformationSummary.FirstOrDefault(l => l.Item2.Significance == LogEntrySignificance.SourceSiteUrl);
            var targetSite = transformationSummary.FirstOrDefault(l => l.Item2.Significance == LogEntrySignificance.TargetSiteUrl);
            var assetsTransferred = transformationSummary.Where(l => l.Item2.Significance == LogEntrySignificance.AssetTransferred);
            var assetsTransferredCount = assetsTransferred.Count();
            var baseTenantUrl = "";

            try
            {
                if (sourceSite != default(Tuple<LogLevel, LogEntry>) && sourceSite.Item2.Message.ContainsIgnoringCasing("https://"))
                {
                    Uri siteUri = new Uri(sourceSite?.Item2.Message);
                    string host = $"{siteUri.Scheme}://{siteUri.DnsSafeHost}";
                    baseTenantUrl = host;
                }
            }
            catch (Exception)
            {
                //Swallow
            }

            #region Calculate Span from Log Timings

            var spanResult = "";
            var logStart = allLogs.FirstOrDefault();
            var logEnd = allLogs.LastOrDefault();

            if (logStart != default(Tuple<LogLevel, LogEntry>) && logEnd != default(Tuple<LogLevel, LogEntry>))
            {
                TimeSpan span = logEnd.Item2.EntryTime.Subtract(logStart.Item2.EntryTime);
                spanResult = string.Format("{0:D2}:{1:D2}:{2:D2}", span.Hours, span.Minutes, span.Seconds);

            }

            #endregion

            //Details
            var logDetails = allLogs.Where(l => l.Item2.Heading != LogStrings.Heading_PageTransformationInfomation &&
                                                l.Item2.Heading != LogStrings.Heading_Summary);

            var pageId = allLogs.First()?.Item2.PageId;

            var logErrors = logDetails.Where(l => l.Item1 == LogLevel.Error);
            var logErrorCount = logErrors.Count();
            var logWarnings = logDetails.Where(l => l.Item1 == LogLevel.Warning);
            var logWarningsCount = logWarnings.Count();

            // Report Content
            if (firstRun)
            {
                report.AppendLine($"{Heading1} {LogStrings.Report_ModernisationReport}");
                report.AppendLine();
            }

            if (!includeVerbose)
            {
                //Summary details only

                //Fields we need: 
                // PageID, Source Page, Date, Site, Duration, Cross-site transfer mode, Target Page Url, Number of Warnings, Number of Errors

                if (firstRun)
                {
                    report.AppendLine($"Source Page {TableColumnSeperator} Date {TableColumnSeperator} Duration {TableColumnSeperator} Target Page Url {TableColumnSeperator} No. of Warnings {TableColumnSeperator} No. of Errors");

                    report.AppendLine($"{TableHeaderColumn} {TableColumnSeperator} {TableHeaderColumn} {TableColumnSeperator} {TableHeaderColumn} {TableColumnSeperator} {TableHeaderColumn} {TableColumnSeperator} {TableHeaderColumn} {TableColumnSeperator} {TableHeaderColumn} {TableColumnSeperator} {TableHeaderColumn} {TableColumnSeperator} {TableHeaderColumn}");

                }

                report.AppendLine($"[{sourcePage?.Item2.Message}]({baseTenantUrl}{sourcePage?.Item2.Message}) {TableColumnSeperator} {reportDate} {TableColumnSeperator} {spanResult} {TableColumnSeperator} [{targetPage?.Item2.Message}]({baseTenantUrl}{targetPage?.Item2.Message}) {TableColumnSeperator} {logWarningsCount} {TableColumnSeperator} {logErrorCount}");


            }
            else
            {
                // Verbose details

                #region Transform Overview

                report.AppendLine($"{Heading2} {LogStrings.Report_TransformationDetails}");
                report.AppendLine();
                report.AppendLine($"{UnorderedListItem} {LogStrings.Report_ReportDate}: {reportDate}");
                report.AppendLine($"{UnorderedListItem} {LogStrings.Report_TransformDuration}: {spanResult}");

                foreach (var log in transformationSummary)
                {
                    var signifcance = "";
                    switch (log.Item2.Significance)
                    {
                        case LogEntrySignificance.AssetTransferred:
                            signifcance = LogStrings.AssetTransferredToUrl;
                            break;
                        case LogEntrySignificance.SourcePage:
                            signifcance = LogStrings.TransformingPage;
                            break;
                        case LogEntrySignificance.SourceSiteUrl:
                            signifcance = LogStrings.TransformingSite;
                            break;
                        case LogEntrySignificance.TargetPage:
                            signifcance = LogStrings.TransformedPage;
                            break;
                        case LogEntrySignificance.TargetSiteUrl:
                            signifcance = LogStrings.CrossSiteTransferToSite;
                            break;

                    }

                    report.AppendLine($"{UnorderedListItem} {signifcance} {log.Item2.Message}");
                }

                #endregion

                #region Summary Page Transformation Information Settings

                report.AppendLine();
                report.AppendLine($"{Heading3} {LogStrings.Report_TransformationSettings}");
                report.AppendLine();
                report.AppendLine($"{LogStrings.Report_Property} {TableColumnSeperator} {LogStrings.Report_Settings}");
                report.AppendLine($"{TableHeaderColumn} {TableColumnSeperator} {TableHeaderColumn}");

                var transformationSettings = allLogs.Where(l => l.Item2.Heading == LogStrings.Heading_PageTransformationInfomation);
                foreach (var log in transformationSettings)
                {
                    var keyValue = log.Item2.Message.Split(new string[] { LogStrings.KeyValueSeperatorToken }, StringSplitOptions.None);
                    if (keyValue.Length == 2) //Protect output
                    {
                        report.AppendLine($"{keyValue[0] ?? ""} {TableColumnSeperator} {keyValue[1] ?? LogStrings.Report_ValueNotSet}");
                    }
                }

                #endregion

                #region Transformation Operation Details

                report.AppendLine($"{Heading2} {LogStrings.Report_TransformDetails}");
                report.AppendLine();

                report.AppendLine(string.Format(LogStrings.Report_TransformDetailsTableHeader, TableColumnSeperator));
                report.AppendLine($"{TableHeaderColumn} {TableColumnSeperator} {TableHeaderColumn} {TableColumnSeperator} {TableHeaderColumn} ");

                IEnumerable<Tuple<LogLevel, LogEntry>> filteredLogDetails = null;
                if (_includeDebugEntries)
                {
                    filteredLogDetails = logDetails.Where(l => l.Item1 == LogLevel.Debug ||
                                                               l.Item1 == LogLevel.Information ||
                                                               l.Item1 == LogLevel.Warning);
                }
                else
                {
                    filteredLogDetails = logDetails.Where(l => l.Item1 == LogLevel.Information ||
                                                               l.Item1 == LogLevel.Warning);
                }

                foreach (var log in filteredLogDetails)
                {
                    switch (log.Item1)
                    {
                        case LogLevel.Information:
                            report.AppendLine($"{log.Item2.EntryTime} {TableColumnSeperator} {log.Item2.Heading} {TableColumnSeperator} {log.Item2.Message}");
                            break;
                        case LogLevel.Warning:
                            report.AppendLine($"{log.Item2.EntryTime} {TableColumnSeperator} {Bold}{log.Item2.Heading}{Bold} {TableColumnSeperator} {Bold}{log.Item2.Message}{Bold}");
                            break;
                        case LogLevel.Debug:
                            report.AppendLine($"{log.Item2.EntryTime} {TableColumnSeperator} {Italic}{log.Item2.Heading}{Italic} {TableColumnSeperator} {Italic}{log.Item2.Message}{Italic}");
                            break;
                    }
                }

                #endregion

                #region Error Details

                if (logErrorCount > 0)
                {
                    #region Report on Errors

                    report.AppendLine($"{Heading3} {LogStrings.Report_ErrorsOccurred}");
                    report.AppendLine();

                    report.AppendLine(string.Format(LogStrings.Report_TransformErrorsTableHeader, TableColumnSeperator));
                    report.AppendLine($"{TableHeaderColumn} {TableColumnSeperator} {TableHeaderColumn} {TableColumnSeperator} {TableHeaderColumn}");

                    foreach (var log in logErrors)
                    {
                        report.AppendLine($"{log.Item2.EntryTime} {TableColumnSeperator} {log.Item2.Heading} {TableColumnSeperator} {log.Item2.Message}");
                    }

                    #endregion

                }

                #endregion

            }

            return report.ToString();
        }

        /// <summary>
        /// Output the report when flush is called
        /// </summary>
        public virtual void Flush()
        {
            try
            {
                var report = GenerateReport();

                // Dont want to assume locality here
                string logRunTime = _reportDate.ToString().Replace('/', '-').Replace(":", "-").Replace(" ", "-");
                string logFileName = $"Page-Transformation-Report-{logRunTime}{_reportFileName}";

                logFileName = $"{_reportFolder}\\{logFileName}.md";

                using (StreamWriter sw = new StreamWriter(logFileName, true))
                {
                    sw.WriteLine(report);
                }

                // Cleardown all logs
                var logs = _lazyLogInstance.Value;
                logs.RemoveRange(0, logs.Count);

                Console.WriteLine($"Report saved as: {logFileName}");

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error writing to log file: {0} {1}", ex.Message, ex.StackTrace);
            }

        }

    }
}
