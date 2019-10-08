using Microsoft.Graph;
using SharePoint.Scanning.Framework;
using System;
using System.Collections.Generic;

namespace SharePoint.Modernization.Scanner.Results
{
    public class WorkflowScanResult: Scan
    {
        public WorkflowScanResult()
        {
            this.UsedActions = new List<string>();
            this.UnsupportedActionsInFlow = new List<string>();
            this.UsedTriggers = new List<string>();
            this.UnsupportedTriggersInFlow = new List<string>();
            this.LastSubscriptionEdit = DateTime.MinValue;
            this.LastDefinitionEdit = DateTime.MinValue;
        }

        public string ListUrl { get; set; }

        public string ListTitle { get; set; }

        public Guid ListId { get; set; }

        public string ContentTypeName { get; set; }
        public string ContentTypeId { get; set; }

        public Guid DefinitionId { get; set; }

        public Guid SubscriptionId { get; set; }

        /// <summary>
        /// 2010 or 2013 workflow engine
        /// </summary>
        public string Version { get; set; }

        public bool IsOOBWorkflow { get; set; }

        /// <summary>
        /// Site, List, ContentType
        /// </summary>
        public string Scope { get; set; }

        /// <summary>
        /// Site, List or Universal workflow (2013) or N/A (2010)
        /// </summary>
        public string RestrictToType { get; set; }

        public bool Enabled { get; set; }

        public string DefinitionName { get; set; }

        public string DefinitionDescription { get; set; }

        public string SubscriptionName { get; set; }

        public bool HasSubscriptions { get; set; }

        public int ActionCount { get; set; }

        public List<string> UsedActions { get; set; }

        public int ToFLowMappingPercentage { get; set; }

        public List<string> UnsupportedActionsInFlow { get; set; }

        public List<string> UsedTriggers { get; set; }
        public List<string> UnsupportedTriggersInFlow { get; set; }

        public DateTime LastSubscriptionEdit { get; set; }
        public DateTime LastDefinitionEdit { get; set; }
    }
}
