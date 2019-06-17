using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Publishing.Navigation;
using Microsoft.SharePoint.Client.WorkflowServices;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using SharePoint.Modernization.Scanner.Results;
using SharePoint.Scanning.Framework;

namespace SharePoint.Modernization.Scanner.Analyzers
{
    /// <summary>
    /// Workflow analyzer
    /// </summary>
    public class WorkflowAnalyzer: BaseAnalyzer
    {
        #region Construction
        /// <summary>
        /// Workflow analyzer construction
        /// </summary>
        /// <param name="url">Url of the web to be analyzed</param>
        /// <param name="siteColUrl">Url of the site collection hosting this web</param>
        /// <param name="scanJob">Job that launched this analyzer</param>
        public WorkflowAnalyzer(string url, string siteColUrl, ModernizationScanJob scanJob) : base(url, siteColUrl, scanJob)
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
                Web web = cc.Web;

                // Pre-load needed properties in a single call
                cc.Load(web, w => w.Id, w => w.ServerRelativeUrl, w => w.Url, /*w => w.WorkflowTemplates,*/ w => w.WorkflowAssociations);
                web.Context.Load(web, p=>p.Lists.Include(li => li.Id, li => li.Title, li => li.Hidden, li => li.DefaultViewUrl, li => li.BaseTemplate, li => li.RootFolder, li => li.ItemCount, li => li.WorkflowAssociations));
                web.Context.ExecuteQueryRetry();

                var lists = web.Lists;

                // *******************************
                // Site & list level 2013 workflow
                // *******************************

                // Retrieve the 2013 site level workflow definitions (including unpublished ones)
                WorkflowDefinition[] siteDefinitions = null;

                try
                {
                    siteDefinitions = web.GetWorkflowDefinitions(false);
                }
                catch (ServerException)
                {
                    // If there is no workflow service present in the farm this method will throw an error. 
                    // Swallow the exception
                }


                // Retrieve the 2013 site level workflow subscriptions
                WorkflowSubscription[] siteSubscriptions = null;

                try
                {
                    siteSubscriptions = web.GetWorkflowSubscriptions();
                }
                catch (ServerException)
                {
                    // If there is no workflow service present in the farm this method will throw an error. 
                    // Swallow the exception
                }

                // *******************************
                // Site & list level 2010 workflow
                // *******************************

                // Retrieve the 2010 site level workflow associations
                if (web.WorkflowAssociations.Count > 0)
                {
                    foreach(var workflowAssociation in web.WorkflowAssociations)
                    {
                        WorkflowScanResult workflowScanResult = new WorkflowScanResult()
                        {
                            SiteColUrl = this.SiteCollectionUrl,
                            SiteURL = this.SiteUrl,
                            ListTitle = "",
                            ListUrl = "",
                            Version = "2010",
                            Scope = "Site",
                            DefinitionName = workflowAssociation.Name,
                            SubscriptionName = workflowAssociation.Name,
                            HasSubscriptions = true,
                            DefinitionId = Guid.Empty,
                            SubscriptionId = workflowAssociation.Id,
                        };

                        if (!this.ScanJob.WorkflowScanResults.TryAdd($"workflowScanResult.SiteURL.{Guid.NewGuid()}", workflowScanResult))
                        {
                            ScanError error = new ScanError()
                            {
                                Error = $"Could not add workflow scan result for {workflowScanResult.SiteColUrl}",
                                SiteColUrl = this.SiteCollectionUrl,
                                SiteURL = this.SiteUrl,
                                Field1 = "WorkflowAnalyzer",
                            };
                            this.ScanJob.ScanErrors.Push(error);
                        }
                    }
                }

                // We've found SP2013 site scoped workflows
                if (siteDefinitions.Count() > 0)
                {
                    foreach (var siteDefinition in siteDefinitions.Where(p=>p.RestrictToType.Equals("site", StringComparison.InvariantCultureIgnoreCase) || p.RestrictToType.Equals("universal", StringComparison.InvariantCultureIgnoreCase)))
                    {
                        // Check if this workflow is also in use
                        var siteWorkflowSubscriptions = siteSubscriptions.Where(p => p.DefinitionId.Equals(siteDefinition.Id));

                        if (siteWorkflowSubscriptions.Count() > 0)
                        {
                            foreach (var siteWorkflowSubscription in siteWorkflowSubscriptions)
                            {
                                WorkflowScanResult workflowScanResult = new WorkflowScanResult()
                                {
                                    SiteColUrl = this.SiteCollectionUrl,
                                    SiteURL = this.SiteUrl,
                                    ListTitle = "",
                                    ListUrl = "",
                                    Version = "2013",
                                    Scope = siteDefinition.RestrictToType,
                                    DefinitionName = siteDefinition.DisplayName,
                                    SubscriptionName = siteWorkflowSubscription.Name,
                                    HasSubscriptions = true,
                                    DefinitionId = siteDefinition.Id,
                                    SubscriptionId = siteWorkflowSubscription.Id,
                                };

                                if (!this.ScanJob.WorkflowScanResults.TryAdd($"workflowScanResult.SiteURL.{Guid.NewGuid()}", workflowScanResult))
                                {
                                    ScanError error = new ScanError()
                                    {
                                        Error = $"Could not add workflow scan result for {workflowScanResult.SiteColUrl}",
                                        SiteColUrl = this.SiteCollectionUrl,
                                        SiteURL = this.SiteUrl,
                                        Field1 = "WorkflowAnalyzer",
                                    };
                                    this.ScanJob.ScanErrors.Push(error);
                                }
                            }
                        }
                        else
                        {
                            WorkflowScanResult workflowScanResult = new WorkflowScanResult()
                            {
                                SiteColUrl = this.SiteCollectionUrl,
                                SiteURL = this.SiteUrl,
                                ListTitle = "",
                                ListUrl = "",
                                Version = "2013",
                                Scope = "Site",
                                DefinitionName = siteDefinition.DisplayName,
                                SubscriptionName = "",
                                HasSubscriptions = false,
                                DefinitionId = siteDefinition.Id,
                                SubscriptionId = Guid.Empty,
                            };

                            if (!this.ScanJob.WorkflowScanResults.TryAdd($"workflowScanResult.SiteURL.{Guid.NewGuid()}", workflowScanResult))
                            {
                                ScanError error = new ScanError()
                                {
                                    Error = $"Could not add workflow scan result for {workflowScanResult.SiteColUrl}",
                                    SiteColUrl = this.SiteCollectionUrl,
                                    SiteURL = this.SiteUrl,
                                    Field1 = "WorkflowAnalyzer",
                                };
                                this.ScanJob.ScanErrors.Push(error);
                            }
                        }
                    }
                }

                // We've found SP2013 list scoped workflows
                if (siteDefinitions.Count() > 0)
                {
                    foreach (var listDefinition in siteDefinitions.Where(p => p.RestrictToType.Equals("list", StringComparison.InvariantCultureIgnoreCase) || p.RestrictToType.Equals("universal", StringComparison.InvariantCultureIgnoreCase)))
                    {
                        // Check if this workflow is also in use
                        var listWorkflowSubscriptions = siteSubscriptions.Where(p => p.DefinitionId.Equals(listDefinition.Id));

                        if (listWorkflowSubscriptions.Count() > 0)
                        {
                            foreach (var listWorkflowSubscription in listWorkflowSubscriptions)
                            {
                                Guid associatedListId = Guid.Empty;
                                string associatedListTitle = "";
                                string associatedListUrl = "";
                                if (Guid.TryParse(GetWorkflowProperty(listWorkflowSubscription, "Microsoft.SharePoint.ActivationProperties.ListId"), out Guid associatedListIdValue))
                                {
                                    associatedListId = associatedListIdValue;

                                    // Lookup this list and update title and url
                                    var listLookup = lists.Where(p => p.Id.Equals(associatedListId)).FirstOrDefault();
                                    if (listLookup != null)
                                    {
                                        associatedListTitle = listLookup.Title;
                                        associatedListUrl = listLookup.RootFolder.ServerRelativeUrl;
                                    }
                                }

                                WorkflowScanResult workflowScanResult = new WorkflowScanResult()
                                {
                                    SiteColUrl = this.SiteCollectionUrl,
                                    SiteURL = this.SiteUrl,
                                    ListTitle = associatedListTitle,
                                    ListUrl = associatedListUrl,
                                    ListId = associatedListId,
                                    Version = "2013",
                                    Scope = listDefinition.RestrictToType,
                                    DefinitionName = listDefinition.DisplayName,
                                    SubscriptionName = listWorkflowSubscription.Name,
                                    HasSubscriptions = true,
                                    DefinitionId = listDefinition.Id,
                                    SubscriptionId = listWorkflowSubscription.Id,
                                };

                                if (!this.ScanJob.WorkflowScanResults.TryAdd($"workflowScanResult.SiteURL.{Guid.NewGuid()}", workflowScanResult))
                                {
                                    ScanError error = new ScanError()
                                    {
                                        Error = $"Could not add workflow scan result for {workflowScanResult.SiteColUrl}",
                                        SiteColUrl = this.SiteCollectionUrl,
                                        SiteURL = this.SiteUrl,
                                        Field1 = "WorkflowAnalyzer",
                                    };
                                    this.ScanJob.ScanErrors.Push(error);
                                }
                            }
                        }
                        else
                        {
                            WorkflowScanResult workflowScanResult = new WorkflowScanResult()
                            {
                                SiteColUrl = this.SiteCollectionUrl,
                                SiteURL = this.SiteUrl,
                                ListTitle = "",
                                ListUrl = "",
                                ListId = Guid.Empty,
                                Version = "2013",
                                Scope = "List",
                                DefinitionName = listDefinition.DisplayName,
                                SubscriptionName = "",
                                HasSubscriptions = false,
                                DefinitionId = listDefinition.Id,
                                SubscriptionId = Guid.Empty,

                            };

                            if (!this.ScanJob.WorkflowScanResults.TryAdd($"workflowScanResult.SiteURL.{Guid.NewGuid()}", workflowScanResult))
                            {
                                ScanError error = new ScanError()
                                {
                                    Error = $"Could not add workflow scan result for {workflowScanResult.SiteColUrl}",
                                    SiteColUrl = this.SiteCollectionUrl,
                                    SiteURL = this.SiteUrl,
                                    Field1 = "WorkflowAnalyzer",
                                };
                                this.ScanJob.ScanErrors.Push(error);
                            }
                        }
                    }
                }

                foreach (var list in lists)
                {
                    if (list.WorkflowAssociations != null && list.WorkflowAssociations.Count > 0)
                    {
                        foreach (var workflowAssociation in list.WorkflowAssociations)
                        {
                            WorkflowScanResult workflowScanResult = new WorkflowScanResult()
                            {
                                SiteColUrl = this.SiteCollectionUrl,
                                SiteURL = this.SiteUrl,
                                ListTitle = list.Title,
                                ListUrl = list.RootFolder.ServerRelativeUrl,
                                ListId = list.Id,
                                Version = "2010",
                                Scope = "List",
                                DefinitionName = workflowAssociation.Name,
                                SubscriptionName = workflowAssociation.Name,
                                HasSubscriptions = true,
                                DefinitionId = Guid.Empty,
                                SubscriptionId = workflowAssociation.Id,
                            };

                            if (!this.ScanJob.WorkflowScanResults.TryAdd($"workflowScanResult.SiteURL.{Guid.NewGuid()}", workflowScanResult))
                            {
                                ScanError error = new ScanError()
                                {
                                    Error = $"Could not add workflow scan result for {workflowScanResult.SiteColUrl}",
                                    SiteColUrl = this.SiteCollectionUrl,
                                    SiteURL = this.SiteUrl,
                                    Field1 = "WorkflowAnalyzer",
                                };
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

        #region Helper methods
        private string GetWorkflowProperty(WorkflowSubscription subscription, string propertyName)
        {
            if (subscription.PropertyDefinitions.ContainsKey(propertyName))
            {
                return subscription.PropertyDefinitions[propertyName];
            }

            return "";
        }
        #endregion
    }
}
