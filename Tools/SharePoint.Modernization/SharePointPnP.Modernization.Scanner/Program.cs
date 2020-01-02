using SharePoint.Modernization.Scanner.Core;
using SharePoint.Modernization.Scanner.Core.Reports;
using SharePoint.Modernization.Scanner.Core.Telemetry;
using SharePointPnP.Modernization.Scanner.Core;
using SharePointPnP.Modernization.Scanner.Core.Reports;
using System;
using System.Collections.Generic;
using System.IO;

namespace SharePoint.Modernization.Scanner
{
    /// <summary>
    /// SharePoint PnP Modernization scanner
    /// </summary>
    class Program
    {
        private static ScannerTelemetry scannerTelemetry;

        /// <summary>
        /// Main method to execute the program
        /// </summary>
        /// <param name="args">Command line arguments</param>
        [STAThread]
        static void Main(string[] args)
        {
            var options = new Options();

            // Show wizard to help the user with filling the needed scan configuration
            if (args.Length == 0)
            {
                Console.WriteLine("Launching wizard...");
                var wizard = new Forms.Wizard(options);
                wizard.ShowDialog();

                if (string.IsNullOrEmpty(options.User) && string.IsNullOrEmpty(options.ClientID))
                {
                    // Trigger validation which will show usage information
                    options.ValidateOptions(args);
                }
            }
            else
            {
                // Validate commandline options
                options.ValidateOptions(args);
            }

            if (options.ExportPaths != null && options.ExportPaths.Count > 0)
            {
                List<ReportStream> reportStreams = new List<ReportStream>();

                Generator generator = new Generator();

                // populate memory streams for the consolidated groupify report
                foreach (var path in options.ExportPaths)
                {
                    LoadCSVFile(reportStreams, path, Generator.GroupifyCSV);
                    LoadCSVFile(reportStreams, path, Generator.ScannerSummaryCSV);
                }
                // create report
                var groupifyReport = generator.CreateGroupifyReport(reportStreams);
                PersistStream($".\\{Generator.GroupifyReport}", groupifyReport);
                // free memory
                ClearReportStreams(reportStreams);

                // populate memory streams for the consolidated list report
                foreach (var path in options.ExportPaths)
                {
                    LoadCSVFile(reportStreams, path, Generator.ListCSV);
                    LoadCSVFile(reportStreams, path, Generator.ScannerSummaryCSV);
                }
                // create report
                var listReport = generator.CreateListReport(reportStreams);
                PersistStream($".\\{Generator.ListReport}", listReport);
                // free memory
                ClearReportStreams(reportStreams);

                // populate memory streams for the consolidated page report
                foreach (var path in options.ExportPaths)
                {
                    LoadCSVFile(reportStreams, path, Generator.PageCSV);
                    LoadCSVFile(reportStreams, path, Generator.ScannerSummaryCSV);
                }
                // create report
                var pageReport = generator.CreatePageReport(reportStreams);
                PersistStream($".\\{Generator.PageReport}", pageReport);
                // free memory
                ClearReportStreams(reportStreams);

                // populate memory streams for the consolidated publishing report
                foreach (var path in options.ExportPaths)
                {
                    LoadCSVFile(reportStreams, path, Generator.PublishingSiteCSV);
                    LoadCSVFile(reportStreams, path, Generator.PublishingWebCSV);
                    LoadCSVFile(reportStreams, path, Generator.PublishingPageCSV);
                    LoadCSVFile(reportStreams, path, Generator.ScannerSummaryCSV);
                }
                // create report
                var publishingReport = generator.CreatePublishingReport(reportStreams);
                PersistStream($".\\{Generator.PublishingReport}", publishingReport);
                // free memory
                ClearReportStreams(reportStreams);

                // populate memory streams for the consolidated workflow report
                foreach (var path in options.ExportPaths)
                {
                    LoadCSVFile(reportStreams, path, Generator.WorkflowCSV);
                    LoadCSVFile(reportStreams, path, Generator.ScannerSummaryCSV);
                }
                // create report
                var workflowReport = generator.CreateWorkflowReport(reportStreams);
                PersistStream($".\\{Generator.WorkflowReport}", workflowReport);
                // free memory
                ClearReportStreams(reportStreams);

                // populate memory streams for the consolidated InfoPath report
                foreach (var path in options.ExportPaths)
                {
                    LoadCSVFile(reportStreams, path, Generator.InfoPathCSV);
                    LoadCSVFile(reportStreams, path, Generator.ScannerSummaryCSV);
                }
                // create report
                var infoPathReport = generator.CreateInfoPathReport(reportStreams);
                PersistStream($".\\{Generator.InfoPathReport}", infoPathReport);
                // free memory
                ClearReportStreams(reportStreams);

                // populate memory streams for the consolidated blog report
                foreach (var path in options.ExportPaths)
                {
                    LoadCSVFile(reportStreams, path, Generator.BlogWebCSV);
                    LoadCSVFile(reportStreams, path, Generator.BlogPageCSV);
                    LoadCSVFile(reportStreams, path, Generator.ScannerSummaryCSV);
                }
                // create report
                var blogReport = generator.CreateBlogReport(reportStreams);
                PersistStream($".\\{Generator.BlogReport}", blogReport);
                // free memory
                ClearReportStreams(reportStreams);
            }
            else
            {
                try
                {
                    DateTime scanStartDateTime = DateTime.Now;

                    // let's catch unhandled exceptions 
                    AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;

                    string workingFolder = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), DateTime.Now.Ticks.ToString());

                    //Instantiate scan job
                    ModernizationScanJob job = new ModernizationScanJob(options, null, null)
                    {

                        // I'm debugging
                        //UseThreading = false
                    };

                    job.Logger += Job_Logger;

                    scannerTelemetry = job.ScannerTelemetry;

                    job.Execute();

                    // Persist the CSV file streams
                    Directory.CreateDirectory(workingFolder);
                    foreach (var csvStream in job.GeneratedFileStreams)
                    {
                        // Move pointer to start 
                        csvStream.Value.Position = 0;
                        string outputfile = $"{workingFolder}\\{csvStream.Key}";

                        Console.WriteLine("Outputting scan results to {0}", outputfile);
                        using (var fileStream = File.Create(outputfile))
                        {
                            CopyStream(csvStream.Value, fileStream);
                            if (options.SkipReport)
                            {
                                // Close and dispose the stream when we're not generating reports
                                csvStream.Value.Dispose();
                            }
                            else
                            {
                                csvStream.Value.Position = 0;
                            }
                        }
                    }

                    // Create reports
                    if (!options.SkipReport)
                    {
                        List<ReportStream> reportStreams = new List<ReportStream>();

                        foreach (var csvStream in job.GeneratedFileStreams)
                        {
                            reportStreams.Add(new ReportStream()
                            {
                                Name = csvStream.Key,
                                Source = workingFolder,
                                DataStream = csvStream.Value,
                            });
                        }

                        List<string> paths = new List<string>
                        {
                            workingFolder
                        };

                        var generator = new Generator();

                        var groupifyReport = generator.CreateGroupifyReport(reportStreams);
                        PersistStream($"{workingFolder}\\{Generator.GroupifyReport}", groupifyReport);

                        if (Options.IncludeLists(options.Mode))
                        {
                            var listReport = generator.CreateListReport(reportStreams);
                            PersistStream($"{workingFolder}\\{Generator.ListReport}", listReport);
                        }

                        if (Options.IncludePage(options.Mode))
                        {
                            var pageReport = generator.CreatePageReport(reportStreams);
                            PersistStream($"{workingFolder}\\{Generator.PageReport}", pageReport);
                        }

                        if (Options.IncludePublishing(options.Mode))
                        {
                            var publishingReport = generator.CreatePublishingReport(reportStreams);
                            PersistStream($"{workingFolder}\\{Generator.PublishingReport}", publishingReport);
                        }

                        if (Options.IncludeWorkflow(options.Mode))
                        {
                            var workflowReport = generator.CreateWorkflowReport(reportStreams);
                            PersistStream($"{workingFolder}\\{Generator.WorkflowReport}", workflowReport);
                        }

                        if (Options.IncludeInfoPath(options.Mode))
                        {
                            var infoPathReport = generator.CreateInfoPathReport(reportStreams);
                            PersistStream($"{workingFolder}\\{Generator.InfoPathReport}", infoPathReport);
                        }

                        if (Options.IncludeBlog(options.Mode))
                        {
                            var blogReport = generator.CreateBlogReport(reportStreams);
                            PersistStream($"{workingFolder}\\{Generator.BlogReport}", blogReport);
                        }

                        // Dispose streams
                        foreach (var csvStream in job.GeneratedFileStreams)
                        {
                            csvStream.Value.Dispose();
                        }
                    }

                    TimeSpan duration = DateTime.Now.Subtract(scanStartDateTime);
                    if (scannerTelemetry != null)
                    {
                        scannerTelemetry.LogScanDone(duration);
                    }
                }
                finally
                {
                    if (scannerTelemetry != null)
                    {
                        scannerTelemetry.Flush();
                    }
                }
            }            
        }

        private static void Job_Logger(object sender, EventArgs e)
        {
            if ((e as LogEventArgs).Severity == LogSeverity.Error)
            {
                Console.WriteLine($"[Error] {(e as LogEventArgs).Message}");
            }
            else
            {
                Console.WriteLine((e as LogEventArgs).Message);
            }
        }

        private static void ClearReportStreams(List<ReportStream> reportStreams)
        {
            foreach (var reportStream in reportStreams)
            {
                reportStream.DataStream.Dispose();
            }
            reportStreams.Clear();
        }

        private static void LoadCSVFile(List<ReportStream> reportStreams, string path, string fileName)
        {
            string csvFile = Path.Combine(path, fileName);

            if (!File.Exists(csvFile))
            {
                return;
            }

            MemoryStream csvStream = new MemoryStream();
            using (var csvFileStream = File.OpenRead(csvFile))
            {
                csvFileStream.CopyTo(csvStream);
            }

            reportStreams.Add(new ReportStream()
            {
                DataStream = csvStream,
                Name = fileName,
                Source = path,
            });
        }

        private static void PersistStream(string outputfile, Stream excelReport)
        {            
            if (excelReport == null)
            {
                return;
            }

            Console.WriteLine($"Creating {outputfile}");

            using (var fileStream = File.Create(outputfile))
            {
                excelReport.Position = 0;
                CopyStream(excelReport, fileStream);
            }
        }

        private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            if (scannerTelemetry != null)
            {
                scannerTelemetry.LogScanCrash(e.ExceptionObject);
            }
        }

        private static void CopyStream(Stream input, Stream output)
        {
            byte[] buffer = new byte[8 * 1024];
            int len;
            while ((len = input.Read(buffer, 0, buffer.Length)) > 0)
            {
                output.Write(buffer, 0, len);
            }
        }
    }
}
