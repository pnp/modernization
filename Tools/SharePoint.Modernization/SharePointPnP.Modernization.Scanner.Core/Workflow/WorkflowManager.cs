using SharePoint.Modernization.Scanner.Core.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Serialization;

namespace SharePoint.Modernization.Scanner.Core.Workflow
{
    /// <summary>
    /// Class to handle workflow analysis
    /// </summary>
    public sealed class WorkflowManager
    {
        private static readonly Lazy<WorkflowManager> _lazyInstance = new Lazy<WorkflowManager>(() => new WorkflowManager());
        private WorkflowActions defaultWorkflowActions;

        private static readonly string[] SP2013SupportedFlowActions = new string[]
        {
            "Microsoft.SharePoint.WorkflowServices.Activities.Comment",
            "Microsoft.SharePoint.WorkflowServices.Activities.CallHTTPWebService",
            "Microsoft.Activities.BuildDynamicValue",
            "Microsoft.Activities.GetDynamicValueProperty",
            "Microsoft.Activities.CountDynamicValueItems",
            "Microsoft.SharePoint.WorkflowServices.Activities.SetField",
            "System.Activities.Statements.Assign",
            "Microsoft.SharePoint.WorkflowServices.Activities.CreateListItem",
            "Microsoft.SharePoint.WorkflowServices.Activities.UpdateListItem",
            "Microsoft.SharePoint.WorkflowServices.Activities.DeleteListItem",
            "Microsoft.SharePoint.WorkflowServices.Activities.WaitForFieldChange",
            "Microsoft.SharePoint.WorkflowServices.Activities.WaitForItemEvent",
            "Microsoft.SharePoint.WorkflowServices.Activities.CheckOutItem",
            "Microsoft.SharePoint.WorkflowServices.Activities.UndoCheckOutItem",
            "Microsoft.SharePoint.WorkflowServices.Activities.CheckInItem",
            "Microsoft.SharePoint.WorkflowServices.Activities.CopyItem",
            "Microsoft.SharePoint.WorkflowServices.Activities.Email",
            "Microsoft.Activities.Expressions.AddToDate",
            "Microsoft.SharePoint.WorkflowServices.Activities.SetTimeField",
            "Microsoft.SharePoint.WorkflowServices.Activities.DateInterval",
            "Microsoft.SharePoint.WorkflowServices.Activities.ExtractSubstringFromEnd",
            "Microsoft.SharePoint.WorkflowServices.Activities.ExtractSubstringFromStart",
            "Microsoft.SharePoint.WorkflowServices.Activities.ExtractSubstringFromIndex",
            "Microsoft.SharePoint.WorkflowServices.Activities.ExtractSubstringFromIndexLength",
            "Microsoft.Activities.Expressions.Trim",
            "Microsoft.Activities.Expressions.IndexOfString",
            "Microsoft.Activities.Expressions.ReplaceString",
            "Microsoft.SharePoint.WorkflowServices.Activities.DelayFor",
            "Microsoft.SharePoint.WorkflowServices.Activities.DelayUntil",
            "Microsoft.SharePoint.WorkflowServices.Activities.Calc",
            "Microsoft.SharePoint.WorkflowServices.Activities.WriteToHistory",
            "Microsoft.SharePoint.WorkflowServices.Activities.TranslateDocument",
            "Microsoft.SharePoint.WorkflowServices.Activities.SetModerationStatus"
        };

        private static readonly string[] SP2010SupportedFlowActions = new string[]
        {
            "Microsoft.SharePoint.WorkflowActions.EmailActivity",
            "Microsoft.SharePoint.WorkflowActions.WithKey.CollectDataTask",
            "Microsoft.SharePoint.WorkflowActions.TodoItemTask",
            "Microsoft.SharePoint.WorkflowActions.GroupAssignedTask",
            "Microsoft.SharePoint.WorkflowActions.WithKey.SetFieldActivity",
            "Microsoft.SharePoint.WorkflowActions.WithKey.UpdateItemActivity",
            "Microsoft.SharePoint.WorkflowActions.WithKey.CreateItemActivity",
            "Microsoft.SharePoint.WorkflowActions.WithKey.CopyItemActivity",
            "Microsoft.SharePoint.WorkflowActions.WithKey.CheckOutItemActivity",
            "Microsoft.SharePoint.WorkflowActions.WithKey.CheckInItemActivity",
            "Microsoft.SharePoint.WorkflowActions.WithKey.UndoCheckOutItemActivity",
            "Microsoft.SharePoint.WorkflowActions.WithKey.DeleteItemActivity",
            "Microsoft.SharePoint.WorkflowActions.WithKey.WaitForActivity",
            "Microsoft.SharePoint.WorkflowActions.WithKey.WaitForDocumentStatusActivity",
            "Microsoft.SharePoint.WorkflowActions.SetVariableActivity",
            "Microsoft.SharePoint.WorkflowActions.BuildStringActivity",
            "Microsoft.SharePoint.WorkflowActions.MathActivity",
            "Microsoft.SharePoint.WorkflowActions.DelayForActivity",
            "Microsoft.SharePoint.WorkflowActions.DelayUntilActivity",
            "System.Workflow.ComponentModel.TerminateActivity",
            "Microsoft.SharePoint.WorkflowActions.LogToHistoryListActivity",
            "Microsoft.SharePoint.WorkflowActions.WithKey.SetModerationStatusActivity",
            "Microsoft.SharePoint.WorkflowActions.AddTimeToDateActivity",
            "Microsoft.SharePoint.WorkflowActions.SetTimeFieldActivity",
            "Microsoft.SharePoint.WorkflowActions.DateIntervalActivity",
            "Microsoft.SharePoint.WorkflowActions.ExtractSubstringFromEndActivity",
            "Microsoft.SharePoint.WorkflowActions.ExtractSubstringFromStartActivity",
            "Microsoft.SharePoint.WorkflowActions.ExtractSubstringFromIndexActivity",
            "Microsoft.SharePoint.WorkflowActions.ExtractSubstringFromIndexLengthActivity",
            "Microsoft.SharePoint.WorkflowActions.CommentActivity",
            "Microsoft.SharePoint.WorkflowActions.PersistOnCloseActivity",
            "Microsoft.SharePoint.WorkflowActions.WithKey.AddListItemPermissionsActivity",
            "Microsoft.SharePoint.WorkflowActions.WithKey.RemoveListItemPermissionsActivity",
            "Microsoft.SharePoint.WorkflowActions.WithKey.ReplaceListItemPermissionsActivity",
            "Microsoft.SharePoint.WorkflowActions.WithKey.InheritListItemParentPermissionsActivity"
        };

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
        /// Translate workflow trigger to a string
        /// </summary>
        /// <param name="onItemCreate">On create was set</param>
        /// <param name="onItemChange">on change wat set</param>
        /// <param name="allowManual">manual execution is allowed</param>
        /// <returns>string representation of the used workflow triggers</returns>
        public WorkflowTriggerAnalysis ParseWorkflowTriggers(bool onItemCreate, bool onItemChange, bool allowManual)
        {
            List<string> triggers = new List<string>();

            if (onItemCreate)
            {
                triggers.Add("OnCreate");
            }

            if (onItemChange)
            {
                triggers.Add("OnChange");
            }

            if (allowManual)
            {
                triggers.Add("Manual");
            }

            return new WorkflowTriggerAnalysis() { WorkflowTriggers = triggers };
        }

        /// <summary>
        /// Analysis a workflow definition and returns the used OOB actions
        /// </summary>
        /// <param name="workflowDefinition">Workflow definition to analyze</param>
        /// <param name="wfType">2010 or 2013 workflow</param>
        /// <returns>List of OOB actions used in the workflow</returns>
        public WorkflowActionAnalysis ParseWorkflowDefinition(string workflowDefinition, WorkflowTypes wfType)
        {
            try
            {
                if (string.IsNullOrEmpty(workflowDefinition))
                {
                    return null;
                }

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
                var namespacePrefix1 = string.Empty;
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
                XmlNodeList ns1Nodes = null;
                if (wfType == WorkflowTypes.SP2010)
                {
                    var xmlns = xmlDoc.FirstChild.Attributes["xmlns:ns1"];
                    if (xmlns != null)
                    {
                        nameSpaceManager.AddNamespace("ns1", xmlns.Value);
                        namespacePrefix1 = "ns1:";

                        ns1Nodes = xmlDoc.SelectNodes($"//{namespacePrefix1}*", nameSpaceManager);
                    }
                }

                // Iterate over the nodes and "identify the OOB activities"
                List<string> usedOOBWorkflowActivities = new List<string>();
                List<string> unsupportedOOBWorkflowActivities = new List<string>();
                int actionCounter = 0;
                int knownActionCounter = 0;
                int unsupportedActionCounter = 0;

                foreach (XmlNode node in nodes)
                {
                    ParseXmlNode(wfType, usedOOBWorkflowActivities, unsupportedOOBWorkflowActivities, ref actionCounter, ref knownActionCounter, ref unsupportedActionCounter, node);
                }


                if (wfType == WorkflowTypes.SP2010 && ns1Nodes != null && ns1Nodes.Count > 0)
                {
                    foreach (XmlNode node in ns1Nodes)
                    {
                        ParseXmlNode(wfType, usedOOBWorkflowActivities, unsupportedOOBWorkflowActivities, ref actionCounter, ref knownActionCounter, ref unsupportedActionCounter, node);
                    }
                }

                return new WorkflowActionAnalysis()
                {
                    WorkflowActions = usedOOBWorkflowActivities,
                    ActionCount = knownActionCounter,
                    UnsupportedActions = unsupportedOOBWorkflowActivities,
                    UnsupportedAccountCount = unsupportedActionCounter
                };
            }
            catch (Exception ex)
            {
                // TODO
                // Eat exception for now
            }

            return null;
        }

        private void ParseXmlNode(WorkflowTypes wfType, List<string> usedOOBWorkflowActivities, List<string> unsupportedOOBWorkflowActivities, ref int actionCounter, ref int knownActionCounter, ref int unsupportedActionCounter, XmlNode node)
        {
            actionCounter++;

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
                knownActionCounter++;
                if (!usedOOBWorkflowActivities.Contains(defaultOOBWorkflowAction.ActionNameShort))
                {
                    usedOOBWorkflowActivities.Add(defaultOOBWorkflowAction.ActionNameShort);
                }


                if (wfType == WorkflowTypes.SP2010)
                {
                    if (!WorkflowManager.SP2010SupportedFlowActions.Contains(defaultOOBWorkflowAction.ActionName))
                    {
                        unsupportedActionCounter++;
                        unsupportedOOBWorkflowActivities.Add(defaultOOBWorkflowAction.ActionNameShort);
                    }
                }
                else if (wfType == WorkflowTypes.SP2013)
                {
                    if (!WorkflowManager.SP2013SupportedFlowActions.Contains(defaultOOBWorkflowAction.ActionName))
                    {
                        unsupportedActionCounter++;
                        unsupportedOOBWorkflowActivities.Add(defaultOOBWorkflowAction.ActionNameShort);
                    }
                }
            }
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
                fileName = "SharePointPnP.Modernization.Scanner.Core.Workflow.sp2010wfmodel.xml";
            }
            else if (wfType == WorkflowTypes.SP2013)
            {
                fileName = "SharePointPnP.Modernization.Scanner.Core.Workflow.sp2013wfmodel.xml";
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
