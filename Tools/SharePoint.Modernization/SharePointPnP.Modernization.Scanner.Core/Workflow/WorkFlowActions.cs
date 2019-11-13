using System.Collections.Generic;

namespace SharePoint.Modernization.Scanner.Core.Workflow
{

    /// <summary>
    /// Workflow types
    /// </summary>
    public enum WorkflowTypes
    {
        SP2010,
        SP2013
    }

    /// <summary>
    /// Workflow action class
    /// </summary>
    public class WorkflowAction
    {
        public string ActionName { get; set; }
        public string ActionNameShort { get; set; }
    }

    /// <summary>
    /// Class holding the default workflow actions for 2010/2013 SharePoint workflow
    /// </summary>
    public class WorkflowActions
    {
        public WorkflowActions()
        {
            SP2010DefaultActions = new List<WorkflowAction>();
            SP2013DefaultActions = new List<WorkflowAction>();
        }

        public List<WorkflowAction> SP2010DefaultActions { get; set; }
        public List<WorkflowAction> SP2013DefaultActions { get; set; }

    }
}
