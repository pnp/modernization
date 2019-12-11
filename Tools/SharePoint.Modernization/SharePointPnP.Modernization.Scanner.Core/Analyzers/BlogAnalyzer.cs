using Microsoft.SharePoint.Client;
using SharePoint.Modernization.Scanner.Core.Results;
using System;
using System.Linq;

namespace SharePoint.Modernization.Scanner.Core.Analyzers
{
    public class BlogAnalyzer : BaseAnalyzer
    {
        #region Variables
        private const string FileRefField = "FileRef";
        private const string FileLeafRefField = "FileLeafRef";
        #endregion

        #region Construction
        /// <summary>
        /// Blog analyzer construction
        /// </summary>
        /// <param name="url">Url of the web to be analyzed</param>
        /// <param name="siteColUrl">Url of the site collection hosting this web</param>
        /// <param name="scanJob">Job that launched this analyzer</param>
        public BlogAnalyzer(string url, string siteColUrl, ModernizationScanJob scanJob) : base(url, siteColUrl, scanJob)
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
                // Is this a blog site
                if (cc.Web.WebTemplate.Equals("BLOG", StringComparison.InvariantCultureIgnoreCase))
                {
                    var web = cc.Web;

                    base.Analyze(cc);

                    BlogWebScanResult blogWebScanResult = new BlogWebScanResult()
                    {
                        SiteColUrl = this.SiteCollectionUrl,
                        SiteURL = this.SiteUrl,
                        WebRelativeUrl = this.SiteUrl.Replace(this.SiteCollectionUrl, ""),
                    };

                    // Log used web template
                    if (web.WebTemplate != null)
                    {
                        blogWebScanResult.WebTemplate = $"{web.WebTemplate}#{web.Configuration}";
                    }

                    // Load additional web properties
                    web.EnsureProperty(p => p.Language);
                    blogWebScanResult.Language = web.Language;

                    // Get the blog page list
                    List blogList = null;
                    var lists = web.GetListsToScan();
                    if (lists != null)
                    {
                        blogList = lists.Where(p => p.BaseTemplate == (int)ListTemplateType.Posts).FirstOrDefault();
                    }

                    // Query the blog posts
                    if (blogList != null)
                    {
                        CamlQuery query = CamlQuery.CreateAllItemsQuery(10000, new string[] { "Title", "Body", "NumComments", "PostCategory", "PublishedDate", "Modified", "Created", "Editor", "Author" });

                        var pages = blogList.GetItems(query);
                        cc.Load(pages);
                        cc.ExecuteQueryRetry();

                        if (pages != null)
                        {
                            blogWebScanResult.BlogPageCount = pages.Count;
                            blogWebScanResult.LastRecentBlogPageChange = blogList.LastItemUserModifiedDate;
                            blogWebScanResult.LastRecentBlogPagePublish = blogList.LastItemUserModifiedDate;

                            foreach (var page in pages)
                            {
                                string pageUrl = null;
                                try
                                {
                                    if (page.FieldValues.ContainsKey(FileRefField) && !String.IsNullOrEmpty(page[FileRefField].ToString()))
                                    {
                                        pageUrl = page[FileRefField].ToString().ToLower();
                                    }
                                    else
                                    {
                                        //skip page
                                        continue;
                                    }

                                    BlogPageScanResult blogPageScanResult = new BlogPageScanResult()
                                    {
                                        SiteColUrl = this.SiteCollectionUrl,
                                        SiteURL = this.SiteUrl,
                                        WebRelativeUrl = blogWebScanResult.WebRelativeUrl,
                                        PageRelativeUrl = pageUrl,
                                    };

                                    if (page.FieldValues.ContainsKey(FileRefField) && !String.IsNullOrEmpty(page[FileRefField].ToString()))
                                    {
                                        blogPageScanResult.PageTitle = page["Title"].ToString();
                                    }                                     

                                    // Add modified information
                                    blogPageScanResult.ModifiedBy = page.LastModifiedBy();
                                    blogPageScanResult.ModifiedAt = page.LastModifiedDateTime();
                                    blogPageScanResult.PublishedDate = page.LastPublishedDateTime();

                                    if (blogPageScanResult.ModifiedAt > blogWebScanResult.LastRecentBlogPageChange)
                                    {
                                        blogWebScanResult.LastRecentBlogPageChange = blogPageScanResult.ModifiedAt;
                                    }

                                    if (blogPageScanResult.PublishedDate > blogWebScanResult.LastRecentBlogPagePublish)
                                    {
                                        blogWebScanResult.LastRecentBlogPagePublish = blogPageScanResult.PublishedDate;
                                    }

                                    if (!this.ScanJob.BlogPageScanResults.TryAdd($"blogScanResult.PageURL.{Guid.NewGuid()}", blogPageScanResult))
                                    {
                                        ScanError error = new ScanError()
                                        {
                                            Error = $"Could not add blog page scan result for {blogPageScanResult.SiteColUrl}",
                                            SiteColUrl = this.SiteCollectionUrl,
                                            SiteURL = this.SiteUrl,
                                            Field1 = "BlogPageAnalyzer",
                                        };
                                        this.ScanJob.ScanErrors.Push(error);
                                    }
                                }
                                catch(Exception ex)
                                {
                                    ScanError error = new ScanError()
                                    {
                                        Error = ex.Message,
                                        SiteColUrl = this.SiteCollectionUrl,
                                        SiteURL = this.SiteUrl,
                                        Field1 = "BlogPageAnalyzer",
                                        Field2 = ex.StackTrace,
                                        Field3 = pageUrl
                                    };
                                    this.ScanJob.ScanErrors.Push(error);
                                }
                            }
                        }
                    }

                    if (!this.ScanJob.BlogWebScanResults.TryAdd($"blogScanResult.WebURL.{Guid.NewGuid()}", blogWebScanResult))
                    {
                        ScanError error = new ScanError()
                        {
                            Error = $"Could not add blog web scan result for {blogWebScanResult.SiteColUrl}",
                            SiteColUrl = this.SiteCollectionUrl,
                            SiteURL = this.SiteUrl,
                            Field1 = "BlogWebAnalyzer",
                        };
                        this.ScanJob.ScanErrors.Push(error);
                    }
                }
            }
            catch(Exception ex)
            {
                ScanError error = new ScanError()
                {
                    Error = ex.Message,
                    SiteColUrl = this.SiteCollectionUrl,
                    SiteURL = this.SiteUrl,
                    Field1 = "BlogPageAnalyzerMainLoop",
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