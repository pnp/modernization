using GenericParsing;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using SharePointPnP.Modernization.Scanner.Core.Reports;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace SharePoint.Modernization.Scanner.Core.Reports
{
    /// <summary>
    /// Generates Excel based reports that make it easier to consume the collected data
    /// </summary>
    public class Generator
    {
        public const string ScannerSummaryCSV = "ScannerSummary.csv";
        // Groupify report variables
        public const string GroupifyCSV = "ModernizationSiteScanResults.csv";
        public const string GroupifyMasterFile = "groupifymaster.xlsx";
        public const string GroupifyReport = "Office 365 Group Connection Readiness.xlsx";
        // List report variables
        public const string ListCSV = "ModernizationListScanResults.csv";
        public const string ListMasterFile = "listmaster.xlsx";
        public const string ListReport = "Office 365 List Readiness.xlsx";
        // Page report variables
        public const string PageCSV = "PageScanResults.csv";
        public const string PageMasterFile = "pagemaster.xlsx";
        public const string PageReport = "Office 365 Page Transformation Readiness.xlsx";
        // Publishing report variables
        public const string PublishingSiteCSV = "ModernizationPublishingSiteScanResults.csv";
        public const string PublishingWebCSV = "ModernizationPublishingWebScanResults.csv";
        public const string PublishingPageCSV = "ModernizationPublishingPageScanResults.csv";
        public const string PublishingMasterFile = "publishingmaster.xlsx";
        public const string PublishingReport = "Office 365 Publishing Portal Transformation Readiness.xlsx";
        // Workflow report variables
        public const string WorkflowCSV = "ModernizationWorkflowScanResults.csv";
        public const string WorkflowMasterFile = "workflowmaster.xlsx";
        public const string WorkflowReport = "Office 365 Classic workflow inventory.xlsx";
        // InfoPath report variables
        public const string InfoPathCSV = "ModernizationInfoPathScanResults.csv";
        public const string InfoPathMasterFile = "infopathmaster.xlsx";
        public const string InfoPathReport = "Office 365 InfoPath inventory.xlsx";
        // Blog report variables
        public const string BlogWebCSV = "ModernizationBlogWebScanResults.csv";
        public const string BlogPageCSV = "ModernizationBlogPageScanResults.csv";
        public const string BlogMasterFile = "blogmaster.xlsx";
        public const string BlogReport = "Office 365 Blog inventory.xlsx";

        /// <summary>
        /// Create the list dashboard
        /// </summary>
        /// <param name="reportStreams">List with available streams</param>
        /// <returns>Stream containing the created report</returns>
        public Stream CreateListReport(List<ReportStream> reportStreams)
        {
            MemoryStream xlsxReportStream = new MemoryStream();
            DataTable blockedListsTable = null;
            ScanSummary scanSummary = null;

            List<string> dataSources = new List<string>();

            foreach (var reportStream in reportStreams)
            {
                if (!dataSources.Contains(reportStream.Source))
                {
                    dataSources.Add(reportStream.Source);
                }
            }

            DateTime dateCreationTime = DateTime.MinValue;
            if (dataSources.Count == 1)
            {
                dateCreationTime = DateTime.Now;
            }

            // import the data and "clean" it
            var fileLoaded = false;
            foreach (var source in dataSources)
            {
                var dataStream = reportStreams.Where(p => p.Source.Equals(source) && p.Name.Equals(ListCSV, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                if (dataStream == null)
                {
                    // Skipping as one does not always have this report 
                    continue;
                }

                fileLoaded = true;

                TextReader reader = new StreamReader(dataStream.DataStream);
                using (GenericParserAdapter parser = new GenericParserAdapter(reader))
                {
                    parser.FirstRowHasHeader = true;
                    parser.MaxBufferSize = 2000000;
                    parser.ColumnDelimiter = DetectUsedDelimiter(dataStream.DataStream);

                    // Read the file                    
                    var baseTable = parser.GetDataTable();

                    // Handle "wrong" column name used in older versions
                    if (baseTable.Columns.Contains("Only blocked by OOB reaons"))
                    {
                        baseTable.Columns["Only blocked by OOB reaons"].ColumnName = "Only blocked by OOB reasons";
                    }

                    // Table 1
                    var blockedListsTable1 = baseTable.Copy();
                    // clean table
                    string[] columnsToKeep = new string[] { "Url", "Site Url", "Site Collection Url", "List Title", "Only blocked by OOB reasons", "Blocked at site level", "Blocked at web level", "Blocked at list level", "List page render type", "List experience", "Blocked by not being able to load Page", "Blocked by view type", "View type", "Blocked by list base template", "List base template", "Blocked by zero or multiple web parts", "Blocked by JSLink", "Blocked by XslLink", "Blocked by Xsl", "Blocked by JSLink field", "Blocked by business data field", "Blocked by task outcome field", "Blocked by publishingField", "Blocked by geo location field", "Blocked by list custom action" };
                    blockedListsTable1 = DropTableColumns(blockedListsTable1, columnsToKeep);

                    if (blockedListsTable == null)
                    {
                        blockedListsTable = blockedListsTable1;
                    }
                    else
                    {
                        blockedListsTable.Merge(blockedListsTable1);
                    }

                    // Read scanner summary data
                    var summaryStream = reportStreams.Where(p => p.Source.Equals(source) && p.Name.Equals(ScannerSummaryCSV, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                    if (summaryStream == null)
                    {
                        throw new Exception($"Stream {ScannerSummaryCSV} is not available.");
                    }
                    var scanSummary1 = DetectScannerSummary(summaryStream.DataStream);

                    if (scanSummary == null)
                    {
                        scanSummary = scanSummary1;
                    }
                    else
                    {
                        MergeScanSummaries(scanSummary, scanSummary1);
                    }
                }
            }

            if (!fileLoaded)
            {
                // nothing loaded, so nothing to use for report generation
                return null;
            }

            if (blockedListsTable.Rows.Count == 0)
            {
                return null;
            }

            // Get the template Excel file
            using (Stream stream = typeof(Generator).Assembly.GetManifestResourceStream($"SharePointPnP.Modernization.Scanner.Core.Reports.{ListMasterFile}"))
            {
                // Push the data to Excel, starting from an Excel template
                using (var excel = new ExcelPackage(stream))
                {
                    var dashboardSheet = excel.Workbook.Worksheets["Dashboard"];

                    if (scanSummary != null)
                    {
                        if (scanSummary.SiteCollections.HasValue)
                        {
                            dashboardSheet.SetValue("AA7", scanSummary.SiteCollections.Value);
                        }
                        if (scanSummary.Webs.HasValue)
                        {
                            dashboardSheet.SetValue("AC7", scanSummary.Webs.Value);
                        }
                        if (scanSummary.Lists.HasValue)
                        {
                            dashboardSheet.SetValue("AE7", scanSummary.Lists.Value);
                        }
                        if (scanSummary.Duration != null)
                        {
                            dashboardSheet.SetValue("AA8", scanSummary.Duration);
                        }
                        if (scanSummary.Version != null)
                        {
                            dashboardSheet.SetValue("AA9", scanSummary.Version);
                        }
                    }

                    if (dateCreationTime != DateTime.MinValue)
                    {
                        dashboardSheet.SetValue("AA6", dateCreationTime.ToString("G", DateTimeFormatInfo.InvariantInfo));
                    }
                    else
                    {
                        dashboardSheet.SetValue("AA6", "-");
                    }

                    var blockedListsSheet = excel.Workbook.Worksheets["BlockedLists"];
                    InsertTableData(blockedListsSheet.Tables[0], blockedListsTable);

                    // Save the resulting file
                    excel.SaveAs(xlsxReportStream);
                }
            }

            return xlsxReportStream;
        }

        /// <summary>
        /// Create the publishing dashboard
        /// </summary>
        /// <param name="reportStreams">List with available streams</param>
        /// <returns>Stream containing the created report</returns>
        public Stream CreatePublishingReport(List<ReportStream> reportStreams)
        {
            MemoryStream xlsxReportStream = new MemoryStream();
            DataTable pubWebsBaseTable = null;
            DataTable pubPagesBaseTable = null;
            DataTable pubWebsTable = null;
            DataTable pubPagesTable = null;
            DataTable pubPagesMissingWebPartsTable = null;
            ScanSummary scanSummary = null;

            List<string> dataSources = new List<string>();

            foreach (var reportStream in reportStreams)
            {
                if (!dataSources.Contains(reportStream.Source))
                {
                    dataSources.Add(reportStream.Source);
                }
            }

            DateTime dateCreationTime = DateTime.MinValue;
            if (dataSources.Count == 1)
            {
                dateCreationTime = DateTime.Now;
            }

            // import the data and "clean" it
            int publishingSiteCollectionsCount = 0;
            var fileLoaded = false;
            foreach (var source in dataSources)
            {
                var dataStream = reportStreams.Where(p => p.Source.Equals(source) && p.Name.Equals(PublishingWebCSV, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                if (dataStream == null)
                {
                    // Skipping as one does not always have this report 
                    continue;
                }

                fileLoaded = true;

                TextReader reader = new StreamReader(dataStream.DataStream);
                using (GenericParserAdapter parser = new GenericParserAdapter(reader))
                {
                    parser.FirstRowHasHeader = true;
                    parser.MaxBufferSize = 2000000;
                    parser.ColumnDelimiter = DetectUsedDelimiter(dataStream.DataStream);

                    // Read the file                    
                    pubWebsBaseTable = parser.GetDataTable();

                    var pubWebsTable1 = pubWebsBaseTable.Copy();
                    // clean table
                    string[] columnsToKeep = new string[] { "SiteCollectionUrl", "SiteUrl", "WebRelativeUrl", "SiteCollectionComplexity", "WebTemplate", "Level", "PageCount", "Language", "VariationLabels", "VariationSourceLabel", "SiteMasterPage", "SystemMasterPage", "AlternateCSS", "HasIncompatibleUserCustomActions", "AllowedPageLayouts", "PageLayoutsConfiguration", "DefaultPageLayout", "GlobalNavigationType", "GlobalStructuralNavigationShowSubSites", "GlobalStructuralNavigationShowPages", "GlobalStructuralNavigationShowSiblings", "GlobalStructuralNavigationMaxCount", "GlobalManagedNavigationTermSetId", "CurrentNavigationType", "CurrentStructuralNavigationShowSubSites", "CurrentStructuralNavigationShowPages", "CurrentStructuralNavigationShowSiblings", "CurrentStructuralNavigationMaxCount", "CurrentManagedNavigationTermSetId", "ManagedNavigationAddNewPages", "ManagedNavigationCreateFriendlyUrls", "LibraryItemScheduling", "LibraryEnableModeration", "LibraryEnableVersioning", "LibraryEnableMinorVersions", "LibraryApprovalWorkflowDefined", "BrokenPermissionInheritance" };
                    pubWebsTable1 = DropTableColumns(pubWebsTable1, columnsToKeep);

                    if (pubWebsTable == null)
                    {
                        pubWebsTable = pubWebsTable1;
                    }
                    else
                    {
                        pubWebsTable.Merge(pubWebsTable1);
                    }

                    // Read scanner summary data
                    var summaryStream = reportStreams.Where(p => p.Source.Equals(source) && p.Name.Equals(ScannerSummaryCSV, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                    if (summaryStream == null)
                    {
                        throw new Exception($"Stream {ScannerSummaryCSV} is not available.");
                    }
                    var scanSummary1 = DetectScannerSummary(summaryStream.DataStream);

                    if (scanSummary == null)
                    {
                        scanSummary = scanSummary1;
                    }
                    else
                    {
                        MergeScanSummaries(scanSummary, scanSummary1);
                    }
                }

                // Load the site CSV file to count the site collection rows
                var dataStream2 = reportStreams.Where(p => p.Source.Equals(source) && p.Name.Equals(PublishingSiteCSV, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                if (dataStream2 != null)
                {
                    TextReader reader2 = new StreamReader(dataStream2.DataStream);
                    using (GenericParserAdapter parser = new GenericParserAdapter(reader2))
                    {
                        parser.FirstRowHasHeader = true;
                        parser.MaxBufferSize = 2000000;
                        parser.ColumnDelimiter = DetectUsedDelimiter(dataStream2.DataStream);

                        var siteData = parser.GetDataTable();
                        if (siteData != null)
                        {
                            //scanSummary.SiteCollections =  siteData.Rows.Count;
                            publishingSiteCollectionsCount = publishingSiteCollectionsCount + siteData.Rows.Count;
                        }
                    }
                }

                // Load the pages CSV file, if available
                var dataStream3 = reportStreams.Where(p => p.Source.Equals(source) && p.Name.Equals(PublishingPageCSV, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                if (dataStream3 != null)
                {
                    TextReader reader3 = new StreamReader(dataStream3.DataStream);
                    using (GenericParserAdapter parser = new GenericParserAdapter(reader3))
                    {
                        parser.FirstRowHasHeader = true;
                        parser.MaxBufferSize = 2000000;
                        parser.ColumnDelimiter = DetectUsedDelimiter(dataStream3.DataStream);

                        // Read the file                    
                        pubPagesBaseTable = parser.GetDataTable();

                        // Table 1
                        var pubPagesTable1 = pubPagesBaseTable.Copy();
                        // clean table
                        string[] columnsToKeep = new string[] { "SiteCollectionUrl", "SiteUrl", "WebRelativeUrl", "PageRelativeUrl", "PageName", "ContentType", "ContentTypeId", "PageLayout", "PageLayoutFile", "PageLayoutWasCustomized", "GlobalAudiences", "SecurityGroupAudiences", "SharePointGroupAudiences", "ModifiedAt", "Mapping %" };
                        pubPagesTable1 = DropTableColumns(pubPagesTable1, columnsToKeep);

                        if (pubPagesTable == null)
                        {
                            pubPagesTable = pubPagesTable1;
                        }
                        else
                        {
                            pubPagesTable.Merge(pubPagesTable1);
                        }

                        // Table 2
                        var pubPagesMissingWebPartsTable1 = pubPagesBaseTable.Copy();
                        // clean table
                        columnsToKeep = new string[] { "SiteCollectionUrl", "SiteUrl", "WebRelativeUrl", "PageRelativeUrl", "Unmapped web parts" };
                        pubPagesMissingWebPartsTable1 = DropTableColumns(pubPagesMissingWebPartsTable1, columnsToKeep);

                        // expand rows
                        pubPagesMissingWebPartsTable1 = ExpandRows(pubPagesMissingWebPartsTable1, "Unmapped web parts");
                        // delete "unneeded" rows
                        for (int i = pubPagesMissingWebPartsTable1.Rows.Count - 1; i >= 0; i--)
                        {
                            DataRow dr = pubPagesMissingWebPartsTable1.Rows[i];
                            if (string.IsNullOrEmpty(dr["Unmapped web parts"].ToString()))
                            {
                                dr.Delete();
                            }
                        }

                        if (pubPagesMissingWebPartsTable == null)
                        {
                            pubPagesMissingWebPartsTable = pubPagesMissingWebPartsTable1;
                        }
                        else
                        {
                            pubPagesMissingWebPartsTable.Merge(pubPagesMissingWebPartsTable1);
                        }

                    }
                }
            }

            if (!fileLoaded)
            {
                // nothing loaded, so nothing to use for report generation
                return null;
            }

            scanSummary.SiteCollections = publishingSiteCollectionsCount;

            // Get the template Excel file
            using (Stream stream = typeof(Generator).Assembly.GetManifestResourceStream($"SharePointPnP.Modernization.Scanner.Core.Reports.{PublishingMasterFile}"))
            {
                // Push the data to Excel, starting from an Excel template
                using (var excel = new ExcelPackage(stream))
                {

                    var dashboardSheet = excel.Workbook.Worksheets["Dashboard"];
                    if (scanSummary != null)
                    {
                        if (scanSummary.SiteCollections.HasValue)
                        {
                            dashboardSheet.SetValue("U7", scanSummary.SiteCollections.Value);
                        }
                        if (scanSummary.Duration != null)
                        {
                            dashboardSheet.SetValue("U8", scanSummary.Duration);
                        }
                        if (scanSummary.Version != null)
                        {
                            dashboardSheet.SetValue("U9", scanSummary.Version);
                        }
                    }

                    if (dateCreationTime > DateTime.Now.Subtract(new TimeSpan(5 * 365, 0, 0, 0, 0)))
                    {
                        dashboardSheet.SetValue("U6", dateCreationTime.ToString("G", DateTimeFormatInfo.InvariantInfo));
                    }
                    else
                    {
                        dashboardSheet.SetValue("U6", "-");
                    }

                    var pubWebsSheet = excel.Workbook.Worksheets["PubWebs"];
                    InsertTableData(pubWebsSheet.Tables[0], pubWebsTable);

                    var pubPagesSheet = excel.Workbook.Worksheets["PubPages"];
                    if (pubPagesTable != null)
                    {
                        InsertTableData(pubPagesSheet.Tables[0], pubPagesTable);
                    }

                    var unmappedWebPartsSheet = excel.Workbook.Worksheets["UnmappedWebParts"];
                    if (pubPagesMissingWebPartsTable != null)
                    {
                        InsertTableData(unmappedWebPartsSheet.Tables[0], pubPagesMissingWebPartsTable);
                    }

                    // Save the resulting file
                    excel.SaveAs(xlsxReportStream);
                }
            }

            return xlsxReportStream;
        }

        /// <summary>
        /// Create the blog dashboard
        /// </summary>
        /// <param name="reportStreams">List with available streams</param>
        /// <returns>Stream containing the created report</returns>
        public Stream CreateBlogReport(List<ReportStream> reportStreams)
        {
            MemoryStream xlsxReportStream = new MemoryStream();
            DataTable blogWebsBaseTable = null;
            DataTable blogPagesBaseTable = null;
            DataTable blogWebsTable = null;
            DataTable blogPagesTable = null;
            ScanSummary scanSummary = null;

            List<string> dataSources = new List<string>();

            foreach (var reportStream in reportStreams)
            {
                if (!dataSources.Contains(reportStream.Source))
                {
                    dataSources.Add(reportStream.Source);
                }
            }

            DateTime dateCreationTime = DateTime.MinValue;
            if (dataSources.Count == 1)
            {
                dateCreationTime = DateTime.Now;
            }

            // import the data and "clean" it
            var fileLoaded = false;

            foreach (var source in dataSources)
            {
                var dataStream = reportStreams.Where(p => p.Source.Equals(source) && p.Name.Equals(BlogWebCSV, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                if (dataStream == null)
                {
                    // Skipping as one does not always have this report 
                    continue;
                }
                fileLoaded = true;

                TextReader reader = new StreamReader(dataStream.DataStream);
                using (GenericParserAdapter parser = new GenericParserAdapter(reader))
                {
                    parser.FirstRowHasHeader = true;
                    parser.MaxBufferSize = 2000000;
                    parser.ColumnDelimiter = DetectUsedDelimiter(dataStream.DataStream);

                    // Read the file                    
                    blogWebsBaseTable = parser.GetDataTable();

                    var blogWebsTable1 = blogWebsBaseTable.Copy();
                    // clean table
                    string[] columnsToKeep = new string[] { "Site Url", "Site Collection Url", "Web Relative Url", "Blog Type", "Web Template", "Language", "Blog Page Count", "Last blog change date", "Last blog publish date", "Change Year", "Change Quarter", "Change Month" };
                    blogWebsTable1 = DropTableColumns(blogWebsTable1, columnsToKeep);

                    if (blogWebsTable == null)
                    {
                        blogWebsTable = blogWebsTable1;
                    }
                    else
                    {
                        blogWebsTable.Merge(blogWebsTable1);
                    }

                    // Read scanner summary data
                    var summaryStream = reportStreams.Where(p => p.Source.Equals(source) && p.Name.Equals(ScannerSummaryCSV, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                    if (summaryStream == null)
                    {
                        throw new Exception($"Stream {ScannerSummaryCSV} is not available.");
                    }
                    var scanSummary1 = DetectScannerSummary(summaryStream.DataStream);

                    if (scanSummary == null)
                    {
                        scanSummary = scanSummary1;
                    }
                    else
                    {
                        MergeScanSummaries(scanSummary, scanSummary1);
                    }
                }

                // Load the pages CSV file, if available
                var dataStream2 = reportStreams.Where(p => p.Source.Equals(source) && p.Name.Equals(BlogPageCSV, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                if (dataStream2 != null)
                {
                    TextReader reader2 = new StreamReader(dataStream2.DataStream);
                    using (GenericParserAdapter parser = new GenericParserAdapter(reader2))
                    {
                        parser.FirstRowHasHeader = true;
                        parser.MaxBufferSize = 2000000;
                        parser.ColumnDelimiter = DetectUsedDelimiter(dataStream2.DataStream);

                        // Read the file                    
                        blogPagesBaseTable = parser.GetDataTable();

                        // Table 1
                        var blogPagesTable1 = blogPagesBaseTable.Copy();
                        // clean table
                        string[] columnsToKeep = new string[] { "Site Url", "Site Collection Url", "Web Relative Url", "Blog Type", "Page Relative Url", "Page Title", "Modified At", "Modified By", "Published At" };
                        blogPagesTable1 = DropTableColumns(blogPagesTable1, columnsToKeep);

                        if (blogPagesTable == null)
                        {
                            blogPagesTable = blogPagesTable1;
                        }
                        else
                        {
                            blogPagesTable.Merge(blogPagesTable1);
                        }
                    }
                }
            }

            if (!fileLoaded)
            {
                // nothing loaded, so nothing to use for report generation
                return null;
            }

            // Get the template Excel file
            using (Stream stream = typeof(Generator).Assembly.GetManifestResourceStream($"SharePointPnP.Modernization.Scanner.Core.Reports.{BlogMasterFile}"))
            {

                // Push the data to Excel, starting from an Excel template
                using (var excel = new ExcelPackage(stream))
                {

                    var dashboardSheet = excel.Workbook.Worksheets["Dashboard"];
                    if (scanSummary != null)
                    {
                        if (scanSummary.SiteCollections.HasValue)
                        {
                            dashboardSheet.SetValue("R7", scanSummary.SiteCollections.Value);
                        }
                        if (scanSummary.Webs.HasValue)
                        {
                            dashboardSheet.SetValue("T7", scanSummary.Webs.Value);
                        }
                        if (scanSummary.Duration != null)
                        {
                            dashboardSheet.SetValue("R8", scanSummary.Duration);
                        }
                        if (scanSummary.Version != null)
                        {
                            dashboardSheet.SetValue("R9", scanSummary.Version);
                        }
                    }

                    if (dateCreationTime > DateTime.Now.Subtract(new TimeSpan(5 * 365, 0, 0, 0, 0)))
                    {
                        dashboardSheet.SetValue("R6", dateCreationTime.ToString("G", DateTimeFormatInfo.InvariantInfo));
                    }
                    else
                    {
                        dashboardSheet.SetValue("R6", "-");
                    }

                    var blogWebsSheet = excel.Workbook.Worksheets["BlogWebs"];
                    InsertTableData(blogWebsSheet.Tables[0], blogWebsTable);

                    var blogPagesSheet = excel.Workbook.Worksheets["BlogPages"];
                    if (blogPagesTable != null)
                    {
                        InsertTableData(blogPagesSheet.Tables[0], blogPagesTable);
                    }

                    // Save the resulting file
                    excel.SaveAs(xlsxReportStream);
                }
            }

            return xlsxReportStream;
        }

        /// <summary>
        /// Create the site page dashboard
        /// </summary>
        /// <param name="reportStreams">List with available streams</param>
        /// <returns>Stream containing the created report</returns>
        public Stream CreatePageReport(List<ReportStream> reportStreams)
        {
            MemoryStream xlsxReportStream = new MemoryStream();
            DataTable readyForPageTransformationTable = null;
            DataTable unmappedWebPartsTable = null;
            ScanSummary scanSummary = null;

            List<string> dataSources = new List<string>();

            foreach (var reportStream in reportStreams)
            {
                if (!dataSources.Contains(reportStream.Source))
                {
                    dataSources.Add(reportStream.Source);
                }
            }

            DateTime dateCreationTime = DateTime.MinValue;
            if (dataSources.Count == 1)
            {
                dateCreationTime = DateTime.Now;
            }

            // import the data and "clean" it
            var fileLoaded = false;
            foreach (var source in dataSources)
            {
                var dataStream = reportStreams.Where(p => p.Source.Equals(source) && p.Name.Equals(PageCSV, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                if (dataStream == null)
                {
                    // Skipping as one does not always have this report 
                    continue;
                }
                fileLoaded = true;

                TextReader reader = new StreamReader(dataStream.DataStream);
                using (GenericParserAdapter parser = new GenericParserAdapter(reader))
                {
                    parser.FirstRowHasHeader = true;
                    parser.MaxBufferSize = 2000000;
                    parser.ColumnDelimiter = DetectUsedDelimiter(dataStream.DataStream);

                    // Read the file                    
                    var baseTable = parser.GetDataTable();

                    // Table 1
                    var readyForPageTransformationTable1 = baseTable.Copy();
                    // clean table
                    string[] columnsToKeep = new string[] { "SiteUrl", "PageUrl", "HomePage", "Type", "Layout", "Mapping %" };
                    readyForPageTransformationTable1 = DropTableColumns(readyForPageTransformationTable1, columnsToKeep);

                    if (readyForPageTransformationTable == null)
                    {
                        readyForPageTransformationTable = readyForPageTransformationTable1;
                    }
                    else
                    {
                        readyForPageTransformationTable.Merge(readyForPageTransformationTable1);
                    }

                    // Table 2
                    var unmappedWebPartsTable1 = baseTable.Copy();

                    // clean table
                    columnsToKeep = new string[] { "SiteUrl", "PageUrl", "Unmapped web parts" };
                    unmappedWebPartsTable1 = DropTableColumns(unmappedWebPartsTable1, columnsToKeep);
                    // expand rows
                    unmappedWebPartsTable1 = ExpandRows(unmappedWebPartsTable1, "Unmapped web parts");
                    // delete "unneeded" rows
                    for (int i = unmappedWebPartsTable1.Rows.Count - 1; i >= 0; i--)
                    {
                        DataRow dr = unmappedWebPartsTable1.Rows[i];
                        if (string.IsNullOrEmpty(dr["Unmapped web parts"].ToString()))
                        {
                            dr.Delete();
                        }
                    }

                    if (unmappedWebPartsTable == null)
                    {
                        unmappedWebPartsTable = unmappedWebPartsTable1;
                    }
                    else
                    {
                        unmappedWebPartsTable.Merge(unmappedWebPartsTable1);
                    }

                    // Read scanner summary data
                    var summaryStream = reportStreams.Where(p => p.Source.Equals(source) && p.Name.Equals(ScannerSummaryCSV, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                    if (summaryStream == null)
                    {
                        throw new Exception($"Stream {ScannerSummaryCSV} is not available.");
                    }
                    var scanSummary1 = DetectScannerSummary(summaryStream.DataStream);

                    if (scanSummary == null)
                    {
                        scanSummary = scanSummary1;
                    }
                    else
                    {
                        MergeScanSummaries(scanSummary, scanSummary1);
                    }
                }
            }

            if (!fileLoaded)
            {
                // nothing loaded, so nothing to use for report generation
                return null;
            }

            // Get the template Excel file
            using (Stream stream = typeof(Generator).Assembly.GetManifestResourceStream($"SharePointPnP.Modernization.Scanner.Core.Reports.{PageMasterFile}"))
            {

                // Push the data to Excel, starting from an Excel template
                //using (var excel = new ExcelPackage(new FileInfo(PageMasterFile)))
                using (var excel = new ExcelPackage(stream))
                {
                    var dashboardSheet = excel.Workbook.Worksheets["Dashboard"];
                    if (scanSummary != null)
                    {
                        if (scanSummary.SiteCollections.HasValue)
                        {
                            dashboardSheet.SetValue("X7", scanSummary.SiteCollections.Value);
                        }
                        if (scanSummary.Webs.HasValue)
                        {
                            dashboardSheet.SetValue("Z7", scanSummary.Webs.Value);
                        }
                        if (scanSummary.Lists.HasValue)
                        {
                            dashboardSheet.SetValue("AB7", scanSummary.Lists.Value);
                        }
                        if (scanSummary.Duration != null)
                        {
                            dashboardSheet.SetValue("X8", scanSummary.Duration);
                        }
                        if (scanSummary.Version != null)
                        {
                            dashboardSheet.SetValue("X9", scanSummary.Version);
                        }
                    }

                    if (dateCreationTime > DateTime.Now.Subtract(new TimeSpan(5 * 365, 0, 0, 0, 0)))
                    {
                        dashboardSheet.SetValue("X6", dateCreationTime.ToString("G", DateTimeFormatInfo.InvariantInfo));
                    }
                    else
                    {
                        dashboardSheet.SetValue("X6", "-");
                    }

                    var readyForPageTransformationSheet = excel.Workbook.Worksheets["ReadyForPageTransformation"];
                    InsertTableData(readyForPageTransformationSheet.Tables[0], readyForPageTransformationTable);

                    var unmappedWebPartsSheet = excel.Workbook.Worksheets["UnmappedWebParts"];
                    InsertTableData(unmappedWebPartsSheet.Tables[0], unmappedWebPartsTable);

                    // Save the resulting file
                    excel.SaveAs(xlsxReportStream);
                }
            }

            return xlsxReportStream;
        }

        /// <summary>
        /// Create the workflow dashboard
        /// </summary>
        /// <param name="reportStreams">List with available streams</param>
        /// <returns>Stream containing the created report</returns>
        public Stream CreateWorkflowReport(List<ReportStream> reportStreams)
        {
            MemoryStream xlsxReportStream = new MemoryStream();
            DataTable workflowTable = null;
            ScanSummary scanSummary = null;

            List<string> dataSources = new List<string>();

            foreach (var reportStream in reportStreams)
            {
                if (!dataSources.Contains(reportStream.Source))
                {
                    dataSources.Add(reportStream.Source);
                }
            }

            DateTime dateCreationTime = DateTime.MinValue;
            if (dataSources.Count == 1)
            {
                dateCreationTime = DateTime.Now;
            }

            // import the data and "clean" it
            var fileLoaded = false;
            foreach (var source in dataSources)
            {
                var dataStream = reportStreams.Where(p => p.Source.Equals(source) && p.Name.Equals(WorkflowCSV, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                if (dataStream == null)
                {
                    // Skipping as one does not always have this report 
                    continue;
                }

                fileLoaded = true;

                TextReader reader = new StreamReader(dataStream.DataStream);
                using (GenericParserAdapter parser = new GenericParserAdapter(reader))
                {
                    parser.FirstRowHasHeader = true;
                    parser.MaxBufferSize = 2000000;
                    parser.ColumnDelimiter = DetectUsedDelimiter(dataStream.DataStream);

                    // Read the file                    
                    var baseTable = parser.GetDataTable();

                    // Table 1
                    var workflowTable1 = baseTable.Copy();
                    // clean table
                    string[] columnsToKeep = new string[] { "Site Url", "Site Collection Url", "Definition Name", "Migration to Flow recommended", "Version", "Scope", "Has subscriptions", "Enabled", "Is OOB", "List Title", "List Url", "List Id", "ContentType Name", "ContentType Id", "Restricted To", "Definition description", "Definition Id", "Subscription Name", "Subscription Id", "Definition Changed On", "Subscription Changed On", "Action Count", "Used Actions", "Used Triggers", "Flow upgradability", "Incompatible Action Count", "Incompatible Actions", "Change Year", "Change Quarter", "Change Month" };
                    workflowTable1 = DropTableColumns(workflowTable1, columnsToKeep);

                    if (workflowTable == null)
                    {
                        workflowTable = workflowTable1;
                    }
                    else
                    {
                        workflowTable.Merge(workflowTable1);
                    }

                    // Read scanner summary data
                    var summaryStream = reportStreams.Where(p => p.Source.Equals(source) && p.Name.Equals(ScannerSummaryCSV, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                    if (summaryStream == null)
                    {
                        throw new Exception($"Stream {ScannerSummaryCSV} is not available.");
                    }
                    var scanSummary1 = DetectScannerSummary(summaryStream.DataStream);

                    if (scanSummary == null)
                    {
                        scanSummary = scanSummary1;
                    }
                    else
                    {
                        MergeScanSummaries(scanSummary, scanSummary1);
                    }
                }
            }

            if (!fileLoaded)
            {
                // nothing loaded, so nothing to use for report generation
                return null;
            }

            // Get the template Excel file
            using (Stream stream = typeof(Generator).Assembly.GetManifestResourceStream($"SharePointPnP.Modernization.Scanner.Core.Reports.{WorkflowMasterFile}"))
            {

                // Push the data to Excel, starting from an Excel template
                using (var excel = new ExcelPackage(stream))
                {
                    var dashboardSheet = excel.Workbook.Worksheets["Dashboard"];
                    if (scanSummary != null)
                    {
                        if (scanSummary.SiteCollections.HasValue)
                        {
                            dashboardSheet.SetValue("U7", scanSummary.SiteCollections.Value);
                        }
                        if (scanSummary.Webs.HasValue)
                        {
                            dashboardSheet.SetValue("W7", scanSummary.Webs.Value);
                        }
                        if (scanSummary.Lists.HasValue)
                        {
                            dashboardSheet.SetValue("Y7", scanSummary.Lists.Value);
                        }
                        if (scanSummary.Duration != null)
                        {
                            dashboardSheet.SetValue("U8", scanSummary.Duration);
                        }
                        if (scanSummary.Version != null)
                        {
                            dashboardSheet.SetValue("U9", scanSummary.Version);
                        }
                    }

                    if (dateCreationTime > DateTime.Now.Subtract(new TimeSpan(5 * 365, 0, 0, 0, 0)))
                    {
                        dashboardSheet.SetValue("U6", dateCreationTime.ToString("G", DateTimeFormatInfo.InvariantInfo));
                    }
                    else
                    {
                        dashboardSheet.SetValue("U6", "-");
                    }

                    var workflowSheet = excel.Workbook.Worksheets["Workflow"];
                    InsertTableData(workflowSheet.Tables[0], workflowTable);

                    // Save the resulting file
                    excel.SaveAs(xlsxReportStream);
                }
            }

            return xlsxReportStream;
        }

        /// <summary>
        /// Create the InfoPath dashboard
        /// </summary>
        /// <param name="reportStreams"></param>
        /// <returns></returns>
        public Stream CreateInfoPathReport(List<ReportStream> reportStreams)
        {
            MemoryStream xlsxReportStream = new MemoryStream();
            DataTable infoPathTable = null;
            ScanSummary scanSummary = null;

            List<string> dataSources = new List<string>();

            foreach (var reportStream in reportStreams)
            {
                if (!dataSources.Contains(reportStream.Source))
                {
                    dataSources.Add(reportStream.Source);
                }
            }

            DateTime dateCreationTime = DateTime.MinValue;
            if (dataSources.Count == 1)
            {
                dateCreationTime = DateTime.Now;
            }

            bool fileLoaded = false;
            // import the data and "clean" it
            foreach (var source in dataSources)
            {
                var dataStream = reportStreams.Where(p => p.Source.Equals(source) && p.Name.Equals(InfoPathCSV, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                if (dataStream == null)
                {
                    // Skipping as one does not always have this report 
                    continue;
                }

                fileLoaded = true;

                TextReader reader = new StreamReader(dataStream.DataStream);

                using (GenericParserAdapter parser = new GenericParserAdapter(reader))
                {
                    parser.FirstRowHasHeader = true;
                    parser.MaxBufferSize = 2000000;
                    parser.ColumnDelimiter = DetectUsedDelimiter(dataStream.DataStream);

                    // Read the file                    
                    var baseTable = parser.GetDataTable();

                    // Table 1
                    var infoPathTable1 = baseTable.Copy();
                    // clean table
                    string[] columnsToKeep = new string[] { "Site Url", "Site Collection Url", "InfoPath Usage", "Enabled", "Last user modified date", "Item count", "List Title", "List Url", "List Id", "Template", "Change Year", "Change Quarter", "Change Month" };
                    infoPathTable1 = DropTableColumns(infoPathTable1, columnsToKeep);

                    if (infoPathTable == null)
                    {
                        infoPathTable = infoPathTable1;
                    }
                    else
                    {
                        infoPathTable.Merge(infoPathTable1);
                    }

                    // Read scanner summary data
                    var summaryStream = reportStreams.Where(p => p.Source.Equals(source) && p.Name.Equals(ScannerSummaryCSV, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                    if (summaryStream == null)
                    {
                        throw new Exception($"Stream {ScannerSummaryCSV} is not available.");
                    }
                    var scanSummary1 = DetectScannerSummary(summaryStream.DataStream);

                    if (scanSummary == null)
                    {
                        scanSummary = scanSummary1;
                    }
                    else
                    {
                        MergeScanSummaries(scanSummary, scanSummary1);
                    }
                }
            }

            if (!fileLoaded)
            {
                // nothing loaded, so nothing to use for report generation
                return null;
            }

            // Get the template Excel file
            using (Stream stream = typeof(Generator).Assembly.GetManifestResourceStream($"SharePointPnP.Modernization.Scanner.Core.Reports.{InfoPathMasterFile}"))
            {
                // Push the data to Excel, starting from an Excel template
                using (var excel = new ExcelPackage(stream))
                {
                    var dashboardSheet = excel.Workbook.Worksheets["Dashboard"];
                    if (scanSummary != null)
                    {
                        if (scanSummary.SiteCollections.HasValue)
                        {
                            dashboardSheet.SetValue("S7", scanSummary.SiteCollections.Value);
                        }
                        if (scanSummary.Webs.HasValue)
                        {
                            dashboardSheet.SetValue("U7", scanSummary.Webs.Value);
                        }
                        if (scanSummary.Lists.HasValue)
                        {
                            dashboardSheet.SetValue("W7", scanSummary.Lists.Value);
                        }
                        if (scanSummary.Duration != null)
                        {
                            dashboardSheet.SetValue("S8", scanSummary.Duration);
                        }
                        if (scanSummary.Version != null)
                        {
                            dashboardSheet.SetValue("S9", scanSummary.Version);
                        }
                    }

                    if (dateCreationTime > DateTime.Now.Subtract(new TimeSpan(5 * 365, 0, 0, 0, 0)))
                    {
                        dashboardSheet.SetValue("S6", dateCreationTime.ToString("G", DateTimeFormatInfo.InvariantInfo));
                    }
                    else
                    {
                        dashboardSheet.SetValue("S6", "-");
                    }

                    var workflowSheet = excel.Workbook.Worksheets["InfoPath"];
                    InsertTableData(workflowSheet.Tables[0], infoPathTable);

                    // Save the resulting file
                    excel.SaveAs(xlsxReportStream);
                }
            }

            return xlsxReportStream;
        }

        /// <summary>
        /// Create the groupify dashboard
        /// </summary>
        /// <param name="reportStreams">List with available streams</param>
        /// <returns>Stream containing the created report</returns>
        public Stream CreateGroupifyReport(List<ReportStream> reportStreams)
        {
            MemoryStream xlsxReportStream = new MemoryStream();
            DataTable readyForGroupifyTable = null;
            DataTable blockersTable = null;
            DataTable warningsTable = null;
            DataTable modernUIWarningsTable = null;
            DataTable permissionWarningsTable = null;
            ScanSummary scanSummary = null;

            List<string> dataSources = new List<string>();

            foreach(var reportStream in reportStreams)
            {
                if (!dataSources.Contains(reportStream.Source))
                {
                    dataSources.Add(reportStream.Source);
                }
            }

            DateTime dateCreationTime = DateTime.MinValue;
            if (dataSources.Count == 1)
            {
                dateCreationTime = DateTime.Now;
            }

            // import the data and "clean" it
            foreach (var source in dataSources)
            {
                var dataStream = reportStreams.Where(p => p.Source.Equals(source) && p.Name.Equals(GroupifyCSV, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                if (dataStream == null)
                {
                    throw new Exception($"Stream for {GroupifyCSV} is not available.");
                }

                TextReader reader = new StreamReader(dataStream.DataStream);
                using (GenericParserAdapter parser = new GenericParserAdapter(reader))
                {
                    parser.FirstRowHasHeader = true;
                    parser.MaxBufferSize = 2000000;
                    parser.ColumnDelimiter = DetectUsedDelimiter(dataStream.DataStream);

                    // Read the file                    
                    var baseTable = parser.GetDataTable();

                    // Table 1: Ready for Groupify
                    var readyForGroupifyTable1 = baseTable.Copy();
                    // clean table
                    string[] columnsToKeep = new string[] { "SiteUrl", "ReadyForGroupify", "GroupMode", "ModernHomePage", "WebTemplate", "HasTeamsTeam", "MasterPage", "AlternateCSS", "UserCustomActions", "SubSites", "SubSitesWithBrokenPermissionInheritance", "ModernPageWebFeatureDisabled", "ModernPageFeatureWasEnabledBySPO", "ModernListSiteBlockingFeatureEnabled", "ModernListWebBlockingFeatureEnabled", "SitePublishingFeatureEnabled", "WebPublishingFeatureEnabled", "Everyone(ExceptExternalUsers)Claim", "UsesADGroups", "ExternalSharing" };
                    readyForGroupifyTable1 = DropTableColumns(readyForGroupifyTable1, columnsToKeep);

                    if (readyForGroupifyTable == null)
                    {
                        readyForGroupifyTable = readyForGroupifyTable1;
                    }
                    else
                    {
                        readyForGroupifyTable.Merge(readyForGroupifyTable1);
                    }

                    // Table 2: Groupify blockers
                    var blockersTable1 = baseTable.Copy();
                    // clean table
                    columnsToKeep = new string[] { "SiteUrl", "ReadyForGroupify", "GroupifyBlockers" };
                    blockersTable1 = DropTableColumns(blockersTable1, columnsToKeep);
                    // expand rows
                    blockersTable1 = ExpandRows(blockersTable1, "GroupifyBlockers");

                    if (blockersTable == null)
                    {
                        blockersTable = blockersTable1;
                    }
                    else
                    {
                        blockersTable.Merge(blockersTable1);
                    }

                    // Table 3: Groupify warnings
                    var warningsTable1 = baseTable.Copy();
                    // clean table
                    columnsToKeep = new string[] { "SiteUrl", "ReadyForGroupify", "GroupifyWarnings" };
                    warningsTable1 = DropTableColumns(warningsTable1, columnsToKeep);
                    // expand rows
                    warningsTable1 = ExpandRows(warningsTable1, "GroupifyWarnings");

                    if (warningsTable == null)
                    {
                        warningsTable = warningsTable1;
                    }
                    else
                    {
                        warningsTable.Merge(warningsTable1);
                    }

                    // Table 4: modern ui warnings
                    var modernUIWarningsTable1 = baseTable.Copy();
                    // clean table
                    columnsToKeep = new string[] { "SiteUrl", "ReadyForGroupify", "ModernUIWarnings" };
                    modernUIWarningsTable1 = DropTableColumns(modernUIWarningsTable1, columnsToKeep);
                    // expand rows
                    modernUIWarningsTable1 = ExpandRows(modernUIWarningsTable1, "ModernUIWarnings");

                    if (modernUIWarningsTable == null)
                    {
                        modernUIWarningsTable = modernUIWarningsTable1;
                    }
                    else
                    {
                        modernUIWarningsTable.Merge(modernUIWarningsTable1);
                    }

                    // Table 5: Groupify warnings
                    var permissionWarningsTable1 = baseTable.Copy();
                    // clean table
                    columnsToKeep = new string[] { "SiteUrl", "ReadyForGroupify", "PermissionWarnings" };
                    permissionWarningsTable1 = DropTableColumns(permissionWarningsTable1, columnsToKeep);
                    // expand rows
                    permissionWarningsTable1 = ExpandRows(permissionWarningsTable1, "PermissionWarnings");

                    if (permissionWarningsTable == null)
                    {
                        permissionWarningsTable = permissionWarningsTable1;
                    }
                    else
                    {
                        permissionWarningsTable.Merge(permissionWarningsTable1);
                    }

                    // Read scanner summary data
                    var summaryStream = reportStreams.Where(p => p.Source.Equals(source) && p.Name.Equals(ScannerSummaryCSV, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
                    if (summaryStream == null)
                    {
                        throw new Exception($"Stream {ScannerSummaryCSV} is not available.");
                    }
                    var scanSummary1 = DetectScannerSummary(summaryStream.DataStream);

                    if (scanSummary == null)
                    {
                        scanSummary = scanSummary1;
                    }
                    else
                    {
                        MergeScanSummaries(scanSummary, scanSummary1);
                    }
                }
            }

            // Get the template Excel file
            using (Stream stream = typeof(Generator).Assembly.GetManifestResourceStream($"SharePointPnP.Modernization.Scanner.Core.Reports.{GroupifyMasterFile}"))
            {
                // Push the data to Excel, starting from an Excel template
                using (var excel = new ExcelPackage(stream))
                {
                    var dashboardSheet = excel.Workbook.Worksheets["Dashboard"];

                    if (scanSummary != null)
                    {
                        if (scanSummary.SiteCollections.HasValue)
                        {
                            dashboardSheet.SetValue("U7", scanSummary.SiteCollections.Value);
                        }
                        if (scanSummary.Webs.HasValue)
                        {
                            dashboardSheet.SetValue("W7", scanSummary.Webs.Value);
                        }
                        if (scanSummary.Lists.HasValue)
                        {
                            dashboardSheet.SetValue("Y7", scanSummary.Lists.Value);
                        }
                        if (scanSummary.Duration != null)
                        {
                            dashboardSheet.SetValue("U8", scanSummary.Duration);
                        }
                        if (scanSummary.Version != null)
                        {
                            dashboardSheet.SetValue("U9", scanSummary.Version);
                        }
                    }

                    if (dateCreationTime > DateTime.Now.Subtract(new TimeSpan(5 * 365, 0, 0, 0, 0)))
                    {
                        dashboardSheet.SetValue("U6", dateCreationTime.ToString("G", DateTimeFormatInfo.InvariantInfo));
                    }
                    else
                    {
                        dashboardSheet.SetValue("U6", "-");
                    }
                    var readyForGroupifySheet = excel.Workbook.Worksheets["ReadyForGroupify"];
                    InsertTableData(readyForGroupifySheet.Tables[0], readyForGroupifyTable);

                    var blockersSheet = excel.Workbook.Worksheets["Blockers"];
                    InsertTableData(blockersSheet.Tables[0], blockersTable);

                    var warningsSheet = excel.Workbook.Worksheets["Warnings"];
                    InsertTableData(warningsSheet.Tables[0], warningsTable);

                    var modernUIWarningsSheet = excel.Workbook.Worksheets["ModernUIWarnings"];
                    InsertTableData(modernUIWarningsSheet.Tables[0], modernUIWarningsTable);

                    var permissionsWarningsSheet = excel.Workbook.Worksheets["PermissionsWarnings"];
                    InsertTableData(permissionsWarningsSheet.Tables[0], permissionWarningsTable);

                    // Save the resulting file
                    excel.SaveAs(xlsxReportStream);
                }
            }

            return xlsxReportStream;
        }

        #region Helper methods
        private void InsertTableData(ExcelTable table, DataTable data)
        {
            // Insert new table data
            var start = table.Address.Start;
            var body = table.WorkSheet.Cells[start.Row + 1, start.Column];

            var outRange = body.LoadFromDataTable(data, false);

            if (outRange != null)
            {
                // Refresh the table ranges so that Excel understands the current size of the table
                var newRange = string.Format("{0}:{1}", start.Address, outRange.End.Address);
                var tableElement = table.TableXml.DocumentElement;
                tableElement.Attributes["ref"].Value = newRange;
                tableElement["autoFilter"].Attributes["ref"].Value = newRange;
            }
        }

        private ScanSummary MergeScanSummaries(ScanSummary baseSummary, ScanSummary summaryToAdd)
        {
            if (summaryToAdd.SiteCollections.HasValue)
            {
                if (baseSummary.SiteCollections.HasValue)
                {                    
                    baseSummary.SiteCollections = baseSummary.SiteCollections.Value + summaryToAdd.SiteCollections.Value;                    
                }
                else
                {
                    baseSummary.SiteCollections = summaryToAdd.SiteCollections.Value;
                }
            }

            if (summaryToAdd.Webs.HasValue)
            {
                if (baseSummary.Webs.HasValue)
                {
                    baseSummary.Webs = baseSummary.Webs.Value + summaryToAdd.Webs.Value;
                }
                else
                {
                    baseSummary.Webs = summaryToAdd.Webs.Value;
                }
            }

            if (summaryToAdd.Lists.HasValue)
            {
                if (baseSummary.Lists.HasValue)
                {
                    baseSummary.Lists = baseSummary.Lists.Value + summaryToAdd.Lists.Value;
                }
                else
                {
                    baseSummary.Lists = summaryToAdd.Lists.Value;
                }
            }

            baseSummary.Duration = "";

            return baseSummary;
        }

        private ScanSummary DetectScannerSummary(Stream dataStreamIn)
        {
            ScanSummary summary = new ScanSummary();

            // Copy the stream because this method is called multiple times and a call closes the stream
            MemoryStream dataStream = new MemoryStream();
            CopyStream(dataStreamIn, dataStream);

            try
            {
                dataStream.Position = 0;
                StreamReader reader = new StreamReader(dataStream);
                using (GenericParserAdapter parser = new GenericParserAdapter(reader))
                {
                    parser.FirstRowHasHeader = true;
                    parser.ColumnDelimiter = DetectUsedDelimiter(dataStream);

                    var baseTable = parser.GetDataTable();
                    List<object> data = new List<object>();
                    if (!string.IsNullOrEmpty(baseTable.Rows[0][0].ToString()) && string.IsNullOrEmpty(baseTable.Rows[0][1].ToString()))
                    {
                        // all might be pushed to first column
                        string[] columns = baseTable.Rows[0][0].ToString().Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                        if (columns.Length < 3)
                        {
                            columns = baseTable.Rows[0][0].ToString().Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
                        }

                        foreach (string column in columns)
                        {
                            data.Add(column);
                        }
                    }
                    else
                    {
                        foreach (DataColumn column in baseTable.Columns)
                        {
                            data.Add(baseTable.Rows[0][column]);
                        }
                    }

                    // move the data into the object instance
                    if (int.TryParse(data[0].ToString(), out int sitecollections))
                    {
                        summary.SiteCollections = sitecollections;
                    }
                    if (int.TryParse(data[1].ToString(), out int webs))
                    {
                        summary.Webs = webs;
                    }
                    if (int.TryParse(data[2].ToString(), out int lists))
                    {
                        summary.Lists = lists;
                    }
                    if (!string.IsNullOrEmpty(data[3].ToString()))
                    {
                        summary.Duration = data[3].ToString();
                    }
                    if (!string.IsNullOrEmpty(data[4].ToString()))
                    {
                        summary.Version = data[4].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                // eat all exceptions here as this is not critical 
            }
            finally
            {
                dataStream.Dispose();
            }

            return summary;
        }

        private DataTable ExpandRows(DataTable table, string column)
        {
            List<DataRow> rowsToAdd = new List<DataRow>();
            foreach (DataRow row in table.Rows)
            {
                if (!string.IsNullOrEmpty(row[column].ToString()) && row[column].ToString().Contains(","))
                {
                    string[] columnToExpand = row[column].ToString().Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);

                    row[column] = columnToExpand[0];

                    for (int i = 1; i < columnToExpand.Length; i++)
                    {
                        var expandedRow = table.NewRow();
                        expandedRow.ItemArray = row.ItemArray;
                        expandedRow[column] = columnToExpand[i];
                        rowsToAdd.Add(expandedRow);
                    }
                }
            }

            foreach (var row in rowsToAdd)
            {
                table.Rows.Add(row);
            }

            return table;
        }

        private DataTable DropTableColumns(DataTable table, string[] columnsToKeep)
        {
            // Get the columns that we don't need
            List<string> toDelete = new List<string>();
            foreach (DataColumn column in table.Columns)
            {
                if (!columnsToKeep.Contains(column.ColumnName))
                {
                    toDelete.Add(column.ColumnName);
                }
            }

            // Delete the unwanted columns
            foreach (var column in toDelete)
            {
                table.Columns.Remove(column);
            }

            // Verify we have all the needed columns
            int i = 0;
            foreach (var column in columnsToKeep)
            {
                if (table.Columns[i].ColumnName == column)
                {
                    i++;
                }
                else
                {
                    throw new Exception($"Required column {column} does not appear in the provided dataset. Did you use a very old version of the scanner or rename/delete columns in the CSV file?");
                }
            }

            return table;
        }

        private char? DetectUsedDelimiter(Stream dataStream)
        {
            try
            {
                dataStream.Position = 0;
                StreamReader reader = new StreamReader(dataStream);

                string line1 = reader.ReadLine() ?? "";

                if (line1.IndexOf(',') > 0)
                {
                    return ',';
                }
                else if (line1.IndexOf(';') > 0)
                {
                    return ';';
                }
                else if (line1.IndexOf('|') > 0)
                {
                    return '|';
                }
                else
                {
                    throw new Exception("CSV file delimiter was not detected");
                }
            }
            finally
            {
                dataStream.Position = 0;
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
        #endregion

    }
}
