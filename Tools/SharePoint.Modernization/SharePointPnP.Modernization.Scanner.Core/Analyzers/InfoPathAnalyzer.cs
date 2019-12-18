using System;
using System.IO;
using System.Linq;
using Microsoft.SharePoint.Client;
using SharePoint.Modernization.Scanner.Core.Results;

namespace SharePoint.Modernization.Scanner.Core.Analyzers
{
    public class InfoPathAnalyzer: BaseAnalyzer
    {
        private static readonly string FormBaseContentType = "0x010101";

        #region Construction
        /// <summary>
        /// InfoPath analyzer construction
        /// </summary>
        /// <param name="url">Url of the web to be analyzed</param>
        /// <param name="siteColUrl">Url of the site collection hosting this web</param>
        /// <param name="scanJob">Job that launched this analyzer</param>
        public InfoPathAnalyzer(string url, string siteColUrl, ModernizationScanJob scanJob) : base(url, siteColUrl, scanJob)
        {            
        }
        #endregion

        #region Analysis
        /// <summary>
        /// Analyses a web for it's workflow usage
        /// </summary>
        /// <param name="cc">ClientContext instance used to retrieve workflow data</param>
        /// <returns>Duration of the workflow analysis</returns>
        public override TimeSpan Analyze(ClientContext cc)
        {
            try
            {
                base.Analyze(cc);

                var baseUri = new Uri(this.SiteUrl);
                var webAppUrl = baseUri.Scheme + "://" + baseUri.Host;

                var lists = cc.Web.GetListsToScan(showHidden: true);

                foreach (var list in lists)
                {
                    if (list.BaseTemplate == (int)ListTemplateType.XMLForm ||
                        (!string.IsNullOrEmpty(list.DocumentTemplateUrl) && list.DocumentTemplateUrl.EndsWith(".xsn", StringComparison.InvariantCultureIgnoreCase))
                       )
                    {
                        // Form libraries depend on InfoPath
                        InfoPathScanResult infoPathScanResult = new InfoPathScanResult()
                        {
                            SiteColUrl = this.SiteCollectionUrl,
                            SiteURL = this.SiteUrl,
                            InfoPathUsage = "FormLibrary",
                            ListTitle = list.Title,
                            ListId = list.Id,
                            ListUrl = list.RootFolder.ServerRelativeUrl,
                            Enabled = true,
                            InfoPathTemplate = !string.IsNullOrEmpty(list.DocumentTemplateUrl) ? Path.GetFileName(list.DocumentTemplateUrl) : "",
                            ItemCount = list.ItemCount,
                            LastItemUserModifiedDate = list.LastItemUserModifiedDate,
                        };

                        if (!this.ScanJob.InfoPathScanResults.TryAdd($"{infoPathScanResult.SiteURL}.{Guid.NewGuid()}", infoPathScanResult))
                        {
                            ScanError error = new ScanError()
                            {
                                Error = $"Could not add formlibrary InfoPath scan result for {infoPathScanResult.SiteColUrl} and list {infoPathScanResult.ListUrl}",
                                SiteColUrl = this.SiteCollectionUrl,
                                SiteURL = this.SiteUrl,
                                Field1 = "InfoPathAnalyzer",
                            };
                            this.ScanJob.ScanErrors.Push(error);
                        }

                    }
                    else if (list.BaseTemplate == (int)ListTemplateType.DocumentLibrary || list.BaseTemplate == (int)ListTemplateType.WebPageLibrary)
                    {
                        // verify if a form content type was attached to this list
                        cc.Load(list, p => p.ContentTypes.Include(c => c.Id, c => c.DocumentTemplateUrl));
                        cc.ExecuteQueryRetry();

                        var formContentTypeFound = list.ContentTypes.Where(c => c.Id.StringValue.StartsWith(FormBaseContentType, StringComparison.InvariantCultureIgnoreCase)).OrderBy(c => c.Id.StringValue.Length).FirstOrDefault();
                        if (formContentTypeFound != null)
                        {
                            // Form libraries depend on InfoPath
                            InfoPathScanResult infoPathScanResult = new InfoPathScanResult()
                            {
                                SiteColUrl = this.SiteCollectionUrl,
                                SiteURL = this.SiteUrl,
                                InfoPathUsage = "ContentType",
                                ListTitle = list.Title,
                                ListId = list.Id,
                                ListUrl = list.RootFolder.ServerRelativeUrl,
                                Enabled = true,
                                InfoPathTemplate = !string.IsNullOrEmpty(formContentTypeFound.DocumentTemplateUrl) ? Path.GetFileName(formContentTypeFound.DocumentTemplateUrl) : "",
                                ItemCount = list.ItemCount,
                                LastItemUserModifiedDate = list.LastItemUserModifiedDate,
                            };

                            if (!this.ScanJob.InfoPathScanResults.TryAdd($"{infoPathScanResult.SiteURL}.{Guid.NewGuid()}", infoPathScanResult))
                            {
                                ScanError error = new ScanError()
                                {
                                    Error = $"Could not add contenttype InfoPath scan result for {infoPathScanResult.SiteColUrl} and list {infoPathScanResult.ListUrl}",
                                    SiteColUrl = this.SiteCollectionUrl,
                                    SiteURL = this.SiteUrl,
                                    Field1 = "InfoPathAnalyzer",
                                };
                                this.ScanJob.ScanErrors.Push(error);
                            }
                        }
                    }
                    else if (list.BaseTemplate == (int)ListTemplateType.GenericList)
                    {
                        try
                        {
                            Folder folder = cc.Web.GetFolderByServerRelativeUrl($"{list.RootFolder.ServerRelativeUrl}/Item");
                            cc.Load(folder, p => p.Properties);
                            cc.ExecuteQueryRetry();

                            if (folder.Properties.FieldValues.ContainsKey("_ipfs_infopathenabled") && folder.Properties.FieldValues.ContainsKey("_ipfs_solutionName"))
                            {
                                bool infoPathEnabled = true;
                                if (bool.TryParse(folder.Properties.FieldValues["_ipfs_infopathenabled"].ToString(), out bool infoPathEnabledParsed))
                                {
                                    infoPathEnabled = infoPathEnabledParsed;
                                }

                                // List with an InfoPath customization
                                InfoPathScanResult infoPathScanResult = new InfoPathScanResult()
                                {
                                    SiteColUrl = this.SiteCollectionUrl,
                                    SiteURL = this.SiteUrl,
                                    InfoPathUsage = "CustomForm",
                                    ListTitle = list.Title,
                                    ListId = list.Id,
                                    ListUrl = list.RootFolder.ServerRelativeUrl,
                                    Enabled = infoPathEnabled,
                                    InfoPathTemplate = folder.Properties.FieldValues["_ipfs_solutionName"].ToString(),
                                    ItemCount = list.ItemCount,
                                    LastItemUserModifiedDate = list.LastItemUserModifiedDate,
                                };

                                if (!this.ScanJob.InfoPathScanResults.TryAdd($"{infoPathScanResult.SiteURL}.{Guid.NewGuid()}", infoPathScanResult))
                                {
                                    ScanError error = new ScanError()
                                    {
                                        Error = $"Could not add customform InfoPath scan result for {infoPathScanResult.SiteColUrl} and list {infoPathScanResult.ListUrl}",
                                        SiteColUrl = this.SiteCollectionUrl,
                                        SiteURL = this.SiteUrl,
                                        Field1 = "InfoPathAnalyzer",
                                    };
                                    this.ScanJob.ScanErrors.Push(error);
                                }
                            }
                        }
                        catch (ServerException ex)
                        {
                            if (((ServerException)ex).ServerErrorTypeName == "System.IO.FileNotFoundException")
                            {
                                // Ignore
                            }
                            else
                            {
                                ScanError error = new ScanError()
                                {
                                    Error = ex.Message,
                                    SiteColUrl = this.SiteCollectionUrl,
                                    SiteURL = this.SiteUrl,
                                    Field1 = "InfoPathAnalyzer",
                                    Field2 = ex.StackTrace,
                                    Field3 = $"{webAppUrl}{list.DefaultViewUrl}"
                                };

                                // Send error to telemetry to make scanner better
                                if (this.ScanJob.ScannerTelemetry != null)
                                {
                                    this.ScanJob.ScannerTelemetry.LogScanError(ex, error);
                                }

                                this.ScanJob.ScanErrors.Push(error);
                            }
                        }
                    }
                }
            }
            finally
            {
                this.StopTime = DateTime.Now;
            }

            // return the duration of this scan
            return new TimeSpan((this.StopTime.Subtract(this.StartTime).Ticks));
        }
        #endregion

    }
}
