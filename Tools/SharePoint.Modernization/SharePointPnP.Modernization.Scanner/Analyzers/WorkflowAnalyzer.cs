using System;
using System.Linq;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Workflow;
using Microsoft.SharePoint.Client.WorkflowServices;
using SharePoint.Modernization.Scanner.Results;
using SharePoint.Scanning.Framework;

namespace SharePoint.Modernization.Scanner.Analyzers
{
    /// <summary>
    /// Workflow analyzer
    /// </summary>
    public class WorkflowAnalyzer: BaseAnalyzer
    {

        private class SP2010WorkFlowAssociation
        {
            public string Scope { get; set; }
            public WorkflowAssociation WorkflowAssociation { get; set; }
            public List AssociatedList { get; set; }
            public ContentType AssociatedContentType { get; set; }
        }

        System.Collections.Generic.List<SP2010WorkFlowAssociation> sp2010WorkflowAssociations;

        #region Construction
        /// <summary>
        /// Workflow analyzer construction
        /// </summary>
        /// <param name="url">Url of the web to be analyzed</param>
        /// <param name="siteColUrl">Url of the site collection hosting this web</param>
        /// <param name="scanJob">Job that launched this analyzer</param>
        public WorkflowAnalyzer(string url, string siteColUrl, ModernizationScanJob scanJob) : base(url, siteColUrl, scanJob)
        {
            this.sp2010WorkflowAssociations = new System.Collections.Generic.List<SP2010WorkFlowAssociation>(20);
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
                cc.Load(web, w => w.Id, w => w.ServerRelativeUrl, w => w.Url, w => w.WorkflowTemplates, w => w.WorkflowAssociations);
                cc.Load(web, p => p.ContentTypes.Include(ct => ct.WorkflowAssociations, ct => ct.Name, ct => ct.StringId));
                cc.Load(web, p=>p.Lists.Include(li => li.Id, li => li.Title, li => li.Hidden, li => li.DefaultViewUrl, li => li.BaseTemplate, li => li.RootFolder, li => li.ItemCount, li => li.WorkflowAssociations));
                cc.ExecuteQueryRetry();

                var lists = web.Lists;

                // *******************************************
                // Site, reusable and list level 2013 workflow
                // *******************************************

                // Retrieve the 2013 site level workflow definitions (including unpublished ones)
                WorkflowDefinition[] siteDefinitions = null;
                // Retrieve the 2013 site level workflow subscriptions
                WorkflowSubscription[] siteSubscriptions = null;

                try
                {
                    var servicesManager = new WorkflowServicesManager(web.Context, web);
                    var deploymentService = servicesManager.GetWorkflowDeploymentService();
                    var subscriptionService = servicesManager.GetWorkflowSubscriptionService();

                    var definitions = deploymentService.EnumerateDefinitions(false);
                    web.Context.Load(definitions);

                    var subscriptions = subscriptionService.EnumerateSubscriptions();
                    web.Context.Load(subscriptions);

                    web.Context.ExecuteQueryRetry();

                    siteDefinitions = definitions.ToArray();
                    siteSubscriptions = subscriptions.ToArray();
                }
                catch (ServerException)
                {
                    // If there is no workflow service present in the farm this method will throw an error. 
                    // Swallow the exception
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
                                    ContentTypeId = "",
                                    ContentTypeName = "",
                                    Version = "2013",
                                    Scope = "Site",
                                    RestrictToType = siteDefinition.RestrictToType,
                                    DefinitionName = siteDefinition.DisplayName,
                                    DefinitionDescription = siteDefinition.Description,
                                    SubscriptionName = siteWorkflowSubscription.Name,
                                    HasSubscriptions = true,
                                    Enabled = siteWorkflowSubscription.Enabled,
                                    DefinitionId = siteDefinition.Id,
                                    SubscriptionId = siteWorkflowSubscription.Id,
                                };

                                if (!this.ScanJob.WorkflowScanResults.TryAdd($"workflowScanResult.SiteURL.{Guid.NewGuid()}", workflowScanResult))
                                {
                                    ScanError error = new ScanError()
                                    {
                                        Error = $"Could not add 2013 site workflow scan result for {workflowScanResult.SiteColUrl}",
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
                                ContentTypeId = "",
                                ContentTypeName = "",
                                Version = "2013",
                                Scope = "Site",
                                RestrictToType = siteDefinition.RestrictToType,
                                DefinitionName = siteDefinition.DisplayName,
                                DefinitionDescription = siteDefinition.Description,
                                SubscriptionName = "",
                                HasSubscriptions = false,
                                Enabled = false,
                                DefinitionId = siteDefinition.Id,
                                SubscriptionId = Guid.Empty,
                            };

                            if (!this.ScanJob.WorkflowScanResults.TryAdd($"workflowScanResult.SiteURL.{Guid.NewGuid()}", workflowScanResult))
                            {
                                ScanError error = new ScanError()
                                {
                                    Error = $"Could not add 2013 site workflow scan result for {workflowScanResult.SiteColUrl}",
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
                                    ContentTypeId = "",
                                    ContentTypeName = "",
                                    Version = "2013",
                                    Scope = "List",
                                    RestrictToType = listDefinition.RestrictToType,
                                    DefinitionName = listDefinition.DisplayName,
                                    DefinitionDescription = listDefinition.Description,
                                    SubscriptionName = listWorkflowSubscription.Name,
                                    HasSubscriptions = true,
                                    Enabled = listWorkflowSubscription.Enabled,
                                    DefinitionId = listDefinition.Id,
                                    SubscriptionId = listWorkflowSubscription.Id,
                                };

                                if (!this.ScanJob.WorkflowScanResults.TryAdd($"workflowScanResult.SiteURL.{Guid.NewGuid()}", workflowScanResult))
                                {
                                    ScanError error = new ScanError()
                                    {
                                        Error = $"Could not add 2013 list workflow scan result for {workflowScanResult.SiteColUrl}",
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
                                ContentTypeId = "",
                                ContentTypeName = "",
                                Version = "2013",
                                Scope = "List",
                                RestrictToType = listDefinition.RestrictToType,
                                DefinitionName = listDefinition.DisplayName,
                                DefinitionDescription = listDefinition.Description,
                                SubscriptionName = "",
                                HasSubscriptions = false,
                                Enabled = false,
                                DefinitionId = listDefinition.Id,
                                SubscriptionId = Guid.Empty,

                            };

                            if (!this.ScanJob.WorkflowScanResults.TryAdd($"workflowScanResult.SiteURL.{Guid.NewGuid()}", workflowScanResult))
                            {
                                ScanError error = new ScanError()
                                {
                                    Error = $"Could not add 2013 list workflow scan result for {workflowScanResult.SiteColUrl}",
                                    SiteColUrl = this.SiteCollectionUrl,
                                    SiteURL = this.SiteUrl,
                                    Field1 = "WorkflowAnalyzer",
                                };
                                this.ScanJob.ScanErrors.Push(error);
                            }
                        }
                    }
                }

                // ***********************************************
                // Site, list and content type level 2010 workflow
                // ***********************************************

                if (web.WorkflowAssociations.Count > 0)
                {
                    foreach (var workflowAssociation in web.WorkflowAssociations)
                    {
                        this.sp2010WorkflowAssociations.Add(new SP2010WorkFlowAssociation() { Scope = "Site", WorkflowAssociation = workflowAssociation });
                    }
                }

                foreach (var list in lists.Where(p => p.WorkflowAssociations.Count > 0))
                {
                    foreach (var workflowAssociation in list.WorkflowAssociations)
                    {
                        this.sp2010WorkflowAssociations.Add(new SP2010WorkFlowAssociation() { Scope = "List", WorkflowAssociation = workflowAssociation, AssociatedList = list });
                    }
                }

                foreach (var ct in web.ContentTypes.Where(p => p.WorkflowAssociations.Count > 0))
                {
                    foreach (var workflowAssociation in ct.WorkflowAssociations)
                    {
                        this.sp2010WorkflowAssociations.Add(new SP2010WorkFlowAssociation() { Scope = "ContentType", WorkflowAssociation = workflowAssociation, AssociatedContentType = ct });
                    }
                }

                // Process 2010 worflows
                if (web.WorkflowTemplates.Count > 0)
                {
                    foreach (var workflowTemplate in web.WorkflowTemplates)
                    {
                        // do we have workflows associated for this template?
                        var associatedWorkflows = this.sp2010WorkflowAssociations.Where(p => p.WorkflowAssociation.BaseId.Equals(workflowTemplate.Id));

                        if (associatedWorkflows.Count() > 0)
                        {
                            foreach(var associatedWorkflow in associatedWorkflows)
                            {
                                WorkflowScanResult workflowScanResult = new WorkflowScanResult()
                                {
                                    SiteColUrl = this.SiteCollectionUrl,
                                    SiteURL = this.SiteUrl,
                                    ListTitle = associatedWorkflow.AssociatedList != null ? associatedWorkflow.AssociatedList.Title : "",
                                    ListUrl = associatedWorkflow.AssociatedList != null ? associatedWorkflow.AssociatedList.RootFolder.ServerRelativeUrl : "",
                                    ListId = associatedWorkflow.AssociatedList != null ? associatedWorkflow.AssociatedList.Id : Guid.Empty,
                                    ContentTypeId = associatedWorkflow.AssociatedContentType != null ? associatedWorkflow.AssociatedContentType.StringId : "",
                                    ContentTypeName = associatedWorkflow.AssociatedContentType != null ? associatedWorkflow.AssociatedContentType.Name : "",
                                    Version = "2010",
                                    Scope = associatedWorkflow.Scope,
                                    RestrictToType = "N/A",
                                    DefinitionName = workflowTemplate.Name,
                                    DefinitionDescription = workflowTemplate.Description,
                                    SubscriptionName = associatedWorkflow.WorkflowAssociation.Name,
                                    HasSubscriptions = true,
                                    Enabled = associatedWorkflow.WorkflowAssociation.Enabled,
                                    DefinitionId = workflowTemplate.Id,
                                    SubscriptionId = associatedWorkflow.WorkflowAssociation.Id,
                                };

                                if (!this.ScanJob.WorkflowScanResults.TryAdd($"workflowScanResult.SiteURL.{Guid.NewGuid()}", workflowScanResult))
                                {
                                    ScanError error = new ScanError()
                                    {
                                        Error = $"Could not add 2010 {associatedWorkflow.Scope} type workflow scan result for {workflowScanResult.SiteColUrl}",
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
                                ContentTypeId = "",
                                ContentTypeName = "",
                                Version = "2010",
                                Scope = "",
                                RestrictToType = "N/A",
                                DefinitionName = workflowTemplate.Name,
                                DefinitionDescription = workflowTemplate.Description,
                                SubscriptionName = "",
                                HasSubscriptions = false,
                                Enabled = false,
                                DefinitionId = workflowTemplate.Id,
                                SubscriptionId = Guid.Empty,
                            };

                            if (!this.ScanJob.WorkflowScanResults.TryAdd($"workflowScanResult.SiteURL.{Guid.NewGuid()}", workflowScanResult))
                            {
                                ScanError error = new ScanError()
                                {
                                    Error = $"Could not add 2010 type workflow scan result for {workflowScanResult.SiteColUrl}",
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
