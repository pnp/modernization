using SharePoint.Scanning.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePoint.Modernization.Scanner.Results
{
    public class WorkflowScanResult: Scan
    {
        public string ListUrl { get; set; }

        public string ListTitle { get; set; }

        public Guid ListId { get; set; }

        public Guid DefinitionId { get; set; }

        public Guid SubscriptionId { get; set; }

        /// <summary>
        /// 2010 or 2013 workflow engine
        /// </summary>
        public string Version { get; set; }

        /// <summary>
        /// Site or List workflow
        /// </summary>
        public string Scope { get; set; }

        public string DefinitionName { get; set; }

        public string SubscriptionName { get; set; }

        public bool HasSubscriptions { get; set; }



    }
}
