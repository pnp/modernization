using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePoint.Modernization.Scanner.Workflow
{
    public class WorkflowTriggerAnalysis
    {
        public WorkflowTriggerAnalysis()
        {
            this.WorkflowTriggers = new List<string>();
            this.UnSupportedTriggers = new List<string>();
        }

        public List<string> WorkflowTriggers { get; set; }
        public List<string> UnSupportedTriggers { get; set; }
    }
}
