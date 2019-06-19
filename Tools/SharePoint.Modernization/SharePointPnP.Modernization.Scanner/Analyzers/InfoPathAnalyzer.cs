using System;
using System.Collections.Generic;
using System.Data.SqlTypes;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using SharePoint.Modernization.Scanner.Results;
using SharePoint.Scanning.Framework;

namespace SharePoint.Modernization.Scanner.Analyzers
{
    public class InfoPathAnalyzer: BaseAnalyzer
    {

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

                var lists = cc.Web.GetListsToScan(showHidden: true);

                foreach (var list in lists)
                {
                    if (list.BaseTemplate == (int)ListTemplateType.XMLForm)
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
                            InfoPathTemplate = ""
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
                        catch(ServerException ex)
                        {
                            if (((ServerException)ex).ServerErrorTypeName == "System.IO.FileNotFoundException")
                            {
                                // Ignore
                            }
                            else
                            {
                                throw;
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
