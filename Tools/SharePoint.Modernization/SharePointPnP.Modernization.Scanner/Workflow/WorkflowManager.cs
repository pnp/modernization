using SharePoint.Modernization.Scanner.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Serialization;

namespace SharePoint.Modernization.Scanner.Workflow
{
    public sealed class WorkflowManager
    {
        private static readonly Lazy<WorkflowManager> _lazyInstance = new Lazy<WorkflowManager>(() => new WorkflowManager());
        private WorkflowActions defaultWorkflowActions;

        /// <summary>
        /// Get's the single workflow manager instance, singleton pattern
        /// </summary>
        public static WorkflowManager Instance
        {
            get
            {
                return _lazyInstance.Value;
            }
        }

        #region Construction
        private WorkflowManager()
        {
            // place for instance initialization code
            defaultWorkflowActions = null;
        }
        #endregion

        /// <summary>
        /// Analysis a workflow definition and returns the used OOB actions
        /// </summary>
        /// <param name="workflowDefinition">Workflow definition to analyze</param>
        /// <param name="wfType">2010 or 2013 workflow</param>
        /// <returns>List of OOB actions used in the workflow</returns>
        public List<string> ParseWorkflowDefinition(string workflowDefinition, WorkflowTypes wfType)
        {           
            try
            {
                var xmlDoc = new XmlDocument();
                xmlDoc.Load(WebpartMappingLoader.GenerateStreamFromString(workflowDefinition));

                //determine  whether document contains namespace
                string namespaceName = "";
                if (wfType == WorkflowTypes.SP2010)
                {
                    namespaceName = "ns0";
                }
                else if (wfType == WorkflowTypes.SP2013)
                {
                    namespaceName = "local";
                }

                var namespacePrefix = string.Empty;
                XmlNamespaceManager nameSpaceManager = null;
                if (xmlDoc.FirstChild.Attributes != null)
                {
                    var xmlns = xmlDoc.FirstChild.Attributes[$"xmlns:{namespaceName}"];
                    if (xmlns != null)
                    {
                        nameSpaceManager = new XmlNamespaceManager(xmlDoc.NameTable);
                        nameSpaceManager.AddNamespace(namespaceName, xmlns.Value);
                        namespacePrefix = namespaceName + ":";
                    }
                }

                // Grab all nodes with the workflow action namespace (ns0/local)
                var nodes = xmlDoc.SelectNodes($"//{namespacePrefix}*", nameSpaceManager);

                // Iterate over the nodes and "identify the OOB activities"
                List<string> usedOOBWorkflowActivities = new List<string>();

                foreach (XmlNode node in nodes)
                {
                    WorkflowAction defaultOOBWorkflowAction = null;

                    if (wfType == WorkflowTypes.SP2010)
                    {
                        defaultOOBWorkflowAction = this.defaultWorkflowActions.SP2010DefaultActions.Where(p => p.ActionNameShort == node.LocalName).FirstOrDefault();
                    }
                    else if (wfType == WorkflowTypes.SP2013)
                    {
                        defaultOOBWorkflowAction = this.defaultWorkflowActions.SP2013DefaultActions.Where(p => p.ActionNameShort == node.LocalName).FirstOrDefault();
                    }

                    if (defaultOOBWorkflowAction != null)
                    {
                        if (!usedOOBWorkflowActivities.Contains(defaultOOBWorkflowAction.ActionNameShort))
                        {
                            usedOOBWorkflowActivities.Add(defaultOOBWorkflowAction.ActionNameShort);
                        }
                    }
                }

                return usedOOBWorkflowActivities;
            }
            catch (Exception ex)
            {
                // TODO
                // Eat exception for now
            }

            return null;
        }

        /// <summary>
        /// Trigger the population of the default workflow actions for 2010/2013 workflows
        /// </summary>
        public void LoadWorkflowDefaultActions()
        {
            WorkflowActions wfActions = new WorkflowActions();

            var sp2010Actions = LoadDefaultActions(WorkflowTypes.SP2010);
            var sp2013Actions = LoadDefaultActions(WorkflowTypes.SP2013);

            foreach(var action in sp2010Actions)
            {
                wfActions.SP2010DefaultActions.Add(new WorkflowAction() { ActionName = action, ActionNameShort = GetShortName(action) });
            }

            foreach (var action in sp2013Actions)
            {
                wfActions.SP2013DefaultActions.Add(new WorkflowAction() { ActionName = action, ActionNameShort = GetShortName(action) });
            }

            this.defaultWorkflowActions = wfActions;
        }

        #region Helper methods
        private string GetShortName(string action)
        {
            if (action.Contains("."))
            {
                return action.Substring(action.LastIndexOf(".") + 1);
            }

            return action;
        }

        private List<string> LoadDefaultActions(WorkflowTypes wfType)
        {
            List<string> wfActionsList = new List<string>();

            string fileName = null;

            if (wfType == WorkflowTypes.SP2010)
            {
                fileName = "SharePoint.Modernization.Scanner.Workflow.sp2010wfmodel.xml";
            }
            else if (wfType == WorkflowTypes.SP2013)
            {
                fileName = "SharePoint.Modernization.Scanner.Workflow.sp2013wfmodel.xml";
            }

            var wfModelString = "";
            using (Stream stream = typeof(WorkflowManager).Assembly.GetManifestResourceStream(fileName))
            {
                using (StreamReader reader = new StreamReader(stream))
                {
                    wfModelString = reader.ReadToEnd();
                }
            }

            if (!string.IsNullOrEmpty(wfModelString))
            {
                if (wfType == WorkflowTypes.SP2010)
                {
                    SP2010.WorkflowInfo wfInformation;
                    using (var stream = WebpartMappingLoader.GenerateStreamFromString(wfModelString))
                    {
                        XmlSerializer xmlWorkflowInformation = new XmlSerializer(typeof(SP2010.WorkflowInfo));
                        wfInformation = (SP2010.WorkflowInfo)xmlWorkflowInformation.Deserialize(stream);
                    }

                    foreach(var wfAction in wfInformation.Actions.Action)
                    {
                        if (!wfActionsList.Contains(wfAction.ClassName))
                        {
                            wfActionsList.Add(wfAction.ClassName);
                        }
                    }
                }
                else if (wfType == WorkflowTypes.SP2013)
                {
                    SP2013.WorkflowInfo wfInformation;
                    using (var stream = WebpartMappingLoader.GenerateStreamFromString(wfModelString))
                    {
                        XmlSerializer xmlWorkflowInformation = new XmlSerializer(typeof(SP2013.WorkflowInfo));
                        wfInformation = (SP2013.WorkflowInfo)xmlWorkflowInformation.Deserialize(stream);

                    }

                    foreach (var wfAction in wfInformation.Actions.Action)
                    {
                        if (!wfActionsList.Contains(wfAction.ClassName))
                        {
                            wfActionsList.Add(wfAction.ClassName);
                        }
                    }
                }
            }

            return wfActionsList;
        }
        #endregion
    }
}
