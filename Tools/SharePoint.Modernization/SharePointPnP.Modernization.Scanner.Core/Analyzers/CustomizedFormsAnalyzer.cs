using Microsoft.SharePoint.Client;
using SharePoint.Modernization.Scanner.Core.Results;
using System;

namespace SharePoint.Modernization.Scanner.Core.Analyzers
{
    public class CustomizedFormsAnalyzer: BaseAnalyzer
    {
        #region Construction
        /// <summary>
        /// CustomizedFormsAnalyzer analyzer construction
        /// </summary>
        /// <param name="url">Url of the web to be analyzed</param>
        /// <param name="siteColUrl">Url of the site collection hosting this web</param>
        /// <param name="scanJob">Job that launched this analyzer</param>
        public CustomizedFormsAnalyzer(string url, string siteColUrl, ModernizationScanJob scanJob) : base(url, siteColUrl, scanJob)
        {
        }
        #endregion

        #region Analysis
        /// <summary>
        /// Analyses a web for it's blog page usage
        /// </summary>
        /// <param name="cc">ClientContext instance used to retrieve blog data</param>
        /// <returns>Duration of the blog analysis</returns>
        public override TimeSpan Analyze(ClientContext cc)
        {
            try
            {
                var web = cc.Web;

                base.Analyze(cc);

                // Load the customized forms per site collection
                cc.Load(cc.Site, p=>p.CustomizedFormsPages);
                cc.ExecuteQueryRetry();

                foreach(var formPage in cc.Site.CustomizedFormsPages)
                {
                    CustomizedFormsScanResult customizedFormsScanResult = new CustomizedFormsScanResult()
                    {
                        SiteColUrl = this.SiteCollectionUrl,
                        SiteURL = this.SiteUrl,
                        //WebRelativeUrl = this.SiteUrl.Replace(this.SiteCollectionUrl, ""),
                        FormType = formPage.formType,
                        Url = formPage.Url,
                        PageId = formPage.pageId,
                        WebpartId = formPage.webpartId,                       
                    };

                    if (!this.ScanJob.CustomizedFormsScanResults.TryAdd($"customizedFormsScanResult.WebURL.{Guid.NewGuid()}", customizedFormsScanResult))
                    {
                        ScanError error = new ScanError()
                        {
                            Error = $"Could not add customized forms scan result for {customizedFormsScanResult.SiteColUrl}",
                            SiteColUrl = this.SiteCollectionUrl,
                            SiteURL = this.SiteUrl,
                            Field1 = "CustomizedFormsAnalyzer",
                        };
                        this.ScanJob.ScanErrors.Push(error);
                    }
                }
            }
            catch (Exception ex)
            {
                ScanError error = new ScanError()
                {
                    Error = ex.Message,
                    SiteColUrl = this.SiteCollectionUrl,
                    SiteURL = this.SiteUrl,
                    Field1 = "CustomizedFormsAnalyzerLoop",
                    Field2 = ex.StackTrace
                };
                this.ScanJob.ScanErrors.Push(error);
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
