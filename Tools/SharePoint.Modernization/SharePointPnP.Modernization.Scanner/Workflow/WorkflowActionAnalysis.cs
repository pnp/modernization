using System.Collections.Generic;

namespace SharePoint.Modernization.Scanner.Workflow
{
    public class WorkflowActionAnalysis
    {

        public WorkflowActionAnalysis()
        {
            this.WorkflowActions = new List<string>();
        }

        public List<string> WorkflowActions { get; set; }
        public int ActionCount { get; set; }
    }
}
