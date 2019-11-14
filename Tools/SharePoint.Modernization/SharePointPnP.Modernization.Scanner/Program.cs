using SharePoint.Modernization.Scanner.Core;
using SharePoint.Modernization.Scanner.Core.Reports;
using SharePoint.Modernization.Scanner.Core.Telemetry;
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
                Generator generator = new Generator();
                //generator.CreateGroupifyReport(options.ExportPaths);
                //generator.CreateListReport(options.ExportPaths);
                //generator.CreatePageReport(options.ExportPaths);
                //generator.CreatePublishingReport(options.ExportPaths);
                //generator.CreateWorkflowReport(options.ExportPaths);
                //generator.CreateInfoPathReport(options.ExportPaths);
                //generator.CreateBlogReport(options.ExportPaths);
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
                    ModernizationScanJob job = new ModernizationScanJob(options)
                    {

                        // I'm debugging
                        //UseThreading = false
                    };

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
