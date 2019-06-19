using AngleSharp.Parser.Html;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using OfficeDevPnP.Core.Pages;
using SharePointPnP.Modernization.Framework.Cache;
using SharePointPnP.Modernization.Framework.Entities;
using SharePointPnP.Modernization.Framework.Extensions;
using SharePointPnP.Modernization.Framework.Telemetry;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Reflection;


namespace SharePointPnP.Modernization.Framework.Transform
{
    /// <summary>
    /// Base page transformator class that contains logic that applies for all page transformations
    /// </summary>
    public abstract class BasePageTransformator: BaseTransform
    {
        internal ClientContext sourceClientContext;
        internal ClientContext targetClientContext;
        internal Stopwatch watch;
        internal const string ExecutionLog = "execution.csv";
        internal PageTransformation pageTransformation;
        internal string version = "undefined";
        internal PageTelemetry pageTelemetry;
        internal bool isRootPage = false;

        #region Helper methods
        internal string GetFieldValue(BaseTransformationInformation baseTransformationInformation, string fieldName)
        {

            if (baseTransformationInformation.SourcePage != null)
            {                
               return baseTransformationInformation.SourcePage[fieldName].ToString();                
            }
            else
            {

                if (baseTransformationInformation.SourceFile != null)
                {
                    var fileServerRelativeUrl = baseTransformationInformation.SourceFile.EnsureProperty(p => p.ServerRelativeUrl);

                    // come up with equivalent field values for the page without listitem (so page living in the root folder of the site)
                    if (fieldName.Equals(Constants.FileRefField))
                    {
                        // e.g. /sites/espctest2/SitePages/demo16.aspx
                        return fileServerRelativeUrl;
                    }
                    else if (fieldName.Equals(Constants.FileDirRefField))
                    {
                        // e.g. /sites/espctest2/SitePages
                        return fileServerRelativeUrl.Replace($"/{System.IO.Path.GetFileName(fileServerRelativeUrl)}", "");

                    }
                    else if (fieldName.Equals(Constants.FileLeafRefField))
                    {
                        // e.g. demo16.aspx
                        return System.IO.Path.GetFileName(fileServerRelativeUrl);
                    }
                }
                return "";
            }
        }

        internal bool FieldExistsAndIsUsed(BaseTransformationInformation baseTransformationInformation, string fieldName)
        {
            if (baseTransformationInformation.SourcePage != null)
            {
                return baseTransformationInformation.SourcePage.FieldExistsAndUsed(fieldName);
            }
            else
            {
                return true;
            }
        }

        internal bool IsRootPage(File file)
        {
            if (file != null)
            {
                return true;
            }

            return false;
        }

        internal void RemoveEmptyTextParts(ClientSidePage targetPage)
        {
            var textParts = targetPage.Controls.Where(p => p.Type == typeof(OfficeDevPnP.Core.Pages.ClientSideText));
            if (textParts != null && textParts.Any())
            {
                HtmlParser parser = new HtmlParser(new HtmlParserOptions() { IsEmbedded = true });

                foreach (var textPart in textParts.ToList())
                {
                    using (var document = parser.Parse(((OfficeDevPnP.Core.Pages.ClientSideText)textPart).Text))
                    {
                        if (document.FirstChild != null && string.IsNullOrEmpty(document.FirstChild.TextContent))
                        {
                            LogInfo(LogStrings.TransformRemovingEmptyWebPart, LogStrings.Heading_RemoveEmptyTextParts);
                            // Drop text part
                            targetPage.Controls.Remove(textPart);
                        }
                    }
                }
            }
        }

        internal void RemoveEmptySectionsAndColumns(ClientSidePage targetPage)
        {
            foreach (var section in targetPage.Sections.ToList())
            {
                // First remove all empty sections
                if (section.Controls.Count == 0)
                {
                    targetPage.Sections.Remove(section);
                }
            }

            // Remove empty columns
            foreach (var section in targetPage.Sections)
            {
                if (section.Type == CanvasSectionTemplate.TwoColumn ||
                    section.Type == CanvasSectionTemplate.TwoColumnLeft ||
                    section.Type == CanvasSectionTemplate.TwoColumnRight)
                {
                    var emptyColumn = section.Columns.Where(p => p.Controls.Count == 0).FirstOrDefault();
                    if (emptyColumn != null)
                    {
                        // drop the empty column and change to single column section
                        section.Columns.Remove(emptyColumn);
                        section.Type = CanvasSectionTemplate.OneColumn;
                        section.Columns.First().ResetColumn(0, 12);
                    }
                }
                else if (section.Type == CanvasSectionTemplate.ThreeColumn)
                {
                    var emptyColumns = section.Columns.Where(p => p.Controls.Count == 0);
                    if (emptyColumns != null)
                    {
                        if (emptyColumns.Any() && emptyColumns.Count() == 2)
                        {
                            // drop the two empty columns and change to single column section
                            foreach (var emptyColumn in emptyColumns.ToList())
                            {
                                section.Columns.Remove(emptyColumn);
                            }
                            section.Type = CanvasSectionTemplate.OneColumn;
                            section.Columns.First().ResetColumn(0, 12);
                        }
                        else if (emptyColumns.Any() && emptyColumns.Count() == 1)
                        {
                            // Remove the empty column and change to two column section
                            section.Columns.Remove(emptyColumns.First());
                            section.Type = CanvasSectionTemplate.TwoColumn;
                            int i = 0;
                            foreach (var column in section.Columns)
                            {
                                column.ResetColumn(i, 6);
                                i++;
                            }
                        }
                    }
                }
            }
        }

        internal void ApplyItemLevelPermissions(bool hasTargetContext, ListItem item, ListItemPermission lip, bool alwaysBreakItemLevelPermissions = false)
        {
            if (lip == null || item == null)
            {
                return;
            }

            // Break permission inheritance on the item if not done yet
            if (alwaysBreakItemLevelPermissions || !item.HasUniqueRoleAssignments)
            {
                item.BreakRoleInheritance(false, false);
                //this.sourceClientContext.ExecuteQueryRetry();
                item.Context.ExecuteQueryRetry();
            }

            if (hasTargetContext)
            {
                // Ensure principals are available in the target site
                Dictionary<string, Principal> targetPrincipals = new Dictionary<string, Principal>(lip.Principals.Count);
                foreach (var principal in lip.Principals)
                {
                    var targetPrincipal = GetPrincipal(this.targetClientContext.Web, principal.Key);
                    if (targetPrincipal != null)
                    {
                        if (!targetPrincipals.ContainsKey(principal.Key))
                        {
                            targetPrincipals.Add(principal.Key, targetPrincipal);
                        }
                    }
                }

                // Assign item level permissions          
                foreach (var roleAssignment in lip.RoleAssignments)
                {
                    if (targetPrincipals.TryGetValue(roleAssignment.Member.LoginName, out Principal principal))
                    {
                        var roleDefinitionBindingCollection = new RoleDefinitionBindingCollection(this.targetClientContext);
                        foreach (var roleDef in roleAssignment.RoleDefinitionBindings)
                        {
                            var targetRoleDef = this.targetClientContext.Web.RoleDefinitions.GetByName(roleDef.Name);
                            if (targetRoleDef != null)
                            {
                                roleDefinitionBindingCollection.Add(targetRoleDef);
                            }
                        }
                        item.RoleAssignments.Add(principal, roleDefinitionBindingCollection);
                    }
                }

                this.targetClientContext.ExecuteQueryRetry();
            }
            else
            {
                // Assign item level permissions
                foreach (var roleAssignment in lip.RoleAssignments)
                {
                    if (lip.Principals.TryGetValue(roleAssignment.Member.LoginName, out Principal principal))
                    {
                        var roleDefinitionBindingCollection = new RoleDefinitionBindingCollection(this.sourceClientContext);
                        foreach (var roleDef in roleAssignment.RoleDefinitionBindings)
                        {
                            roleDefinitionBindingCollection.Add(roleDef);
                        }

                        item.RoleAssignments.Add(principal, roleDefinitionBindingCollection);
                    }
                }

                this.sourceClientContext.ExecuteQueryRetry();
            }

            LogInfo(LogStrings.TransformCopiedItemPermissions, LogStrings.Heading_ApplyItemLevelPermissions);
        }

        internal ListItemPermission GetItemLevelPermissions(bool hasTargetContext, List pagesLibrary, ListItem source, ListItem target)
        {
            ListItemPermission lip = null;

            if (source.HasUniqueRoleAssignments)
            {
                // You need to have the ManagePermissions permission before item level permissions can be copied
                if (pagesLibrary.EffectiveBasePermissions.Has(PermissionKind.ManagePermissions))
                {
                    // Copy the unique permissions from source to target
                    // Get the unique permissions
                    this.sourceClientContext.Load(source, a => a.EffectiveBasePermissions, a => a.RoleAssignments.Include(roleAsg => roleAsg.Member.LoginName,
                        roleAsg => roleAsg.RoleDefinitionBindings.Include(roleDef => roleDef.Name, roleDef => roleDef.Description)));
                    this.sourceClientContext.ExecuteQueryRetry();

                    if (source.EffectiveBasePermissions.Has(PermissionKind.ManagePermissions))
                    {
                        // Load the site groups
                        this.sourceClientContext.Load(this.sourceClientContext.Web.SiteGroups, p => p.Include(g => g.LoginName));

                        // Get target page information
                        if (hasTargetContext)
                        {
                            this.targetClientContext.Load(target, p => p.HasUniqueRoleAssignments, p => p.RoleAssignments);
                            this.targetClientContext.Load(this.targetClientContext.Web, p => p.RoleDefinitions);
                        }
                        else
                        {
                            this.sourceClientContext.Load(target, p => p.HasUniqueRoleAssignments, p => p.RoleAssignments);
                        }

                        this.sourceClientContext.ExecuteQueryRetry();

                        if (hasTargetContext)
                        {
                            this.targetClientContext.ExecuteQueryRetry();
                        }

                        Dictionary<string, Principal> principals = new Dictionary<string, Principal>(10);
                        lip = new ListItemPermission()
                        {
                            RoleAssignments = source.RoleAssignments,
                            Principals = principals
                        };

                        // Apply new permissions
                        foreach (var roleAssignment in source.RoleAssignments)
                        {
                            var principal = GetPrincipal(this.sourceClientContext.Web, roleAssignment.Member.LoginName);
                            if (principal != null)
                            {
                                if (!lip.Principals.ContainsKey(roleAssignment.Member.LoginName))
                                {
                                    lip.Principals.Add(roleAssignment.Member.LoginName, principal);
                                }
                            }
                        }
                    }
                }
            }

            LogInfo(LogStrings.TransformGetItemPermissions, LogStrings.Heading_ApplyItemLevelPermissions);

            return lip;
        }

        internal Principal GetPrincipal(Web web, string principalInput)
        {
            Principal principal = this.sourceClientContext.Web.SiteGroups.FirstOrDefault(g => g.LoginName.Equals(principalInput, StringComparison.OrdinalIgnoreCase));

            if (principal == null)
            {
                if (principalInput.Contains("#ext#"))
                {
                    principal = web.SiteUsers.FirstOrDefault(u => u.LoginName.Equals(principalInput));

                    if (principal == null)
                    {
                        //Skipping external user...
                    }
                }
                else
                {
                    try
                    {
                        principal = web.EnsureUser(principalInput);
                        web.Context.ExecuteQueryRetry();
                    }
                    catch (Exception ex)
                    {
                        //Failed to EnsureUser
                        LogError(LogStrings.Error_GetPrincipalFailedEnsureUser, LogStrings.Heading_GetPrincipal, ex);
                    }
                }
            }

            return principal;
        }

        internal void CopyPageMetadata(PageTransformationInformation pageTransformationInformation, ClientSidePage targetPage, List pagesLibrary)
        {
            var fieldsToCopy = CacheManager.Instance.GetFieldsToCopy(this.sourceClientContext.Web, pagesLibrary);
            if (fieldsToCopy.Count > 0)
            {
                // Load the target page list item
                this.sourceClientContext.Load(targetPage.PageListItem);
                this.sourceClientContext.ExecuteQueryRetry();

                // regular fields
                bool isDirty = false;
                foreach (var fieldToCopy in fieldsToCopy.Where(p => p.FieldType != "TaxonomyFieldTypeMulti" && p.FieldType != "TaxonomyFieldType"))
                {
                    if (pageTransformationInformation.SourcePage[fieldToCopy.FieldName] != null)
                    {
                        targetPage.PageListItem[fieldToCopy.FieldName] = pageTransformationInformation.SourcePage[fieldToCopy.FieldName];
                        isDirty = true;

                        LogInfo($"{LogStrings.TransformCopyingMetaDataField} {fieldToCopy.FieldName}", LogStrings.Heading_CopyingPageMetadata);
                    }
                }

                if (isDirty)
                {
                    targetPage.PageListItem.Update();
                    this.sourceClientContext.Load(targetPage.PageListItem);
                    this.sourceClientContext.ExecuteQueryRetry();
                    isDirty = false;
                }

                // taxonomy fields
                foreach (var fieldToCopy in fieldsToCopy.Where(p => p.FieldType == "TaxonomyFieldTypeMulti" || p.FieldType == "TaxonomyFieldType"))
                {
                    switch (fieldToCopy.FieldType)
                    {
                        case "TaxonomyFieldTypeMulti":
                            {
                                var taxFieldBeforeCast = pagesLibrary.Fields.Where(p => p.Id.Equals(fieldToCopy.FieldId)).FirstOrDefault();
                                if (taxFieldBeforeCast != null)
                                {
                                    var taxField = this.sourceClientContext.CastTo<TaxonomyField>(taxFieldBeforeCast);

                                    if (pageTransformationInformation.SourcePage[fieldToCopy.FieldName] != null)
                                    {
                                        if (pageTransformationInformation.SourcePage[fieldToCopy.FieldName] is TaxonomyFieldValueCollection)
                                        {
                                            var valueCollectionToCopy = (pageTransformationInformation.SourcePage[fieldToCopy.FieldName] as TaxonomyFieldValueCollection);
                                            var taxonomyFieldValueArray = valueCollectionToCopy.Select(taxonomyFieldValue => $"-1;#{taxonomyFieldValue.Label}|{taxonomyFieldValue.TermGuid}");
                                            var valueCollection = new TaxonomyFieldValueCollection(this.sourceClientContext, string.Join(";#", taxonomyFieldValueArray), taxField);
                                            taxField.SetFieldValueByValueCollection(targetPage.PageListItem, valueCollection);
                                        }
                                        else if (pageTransformationInformation.SourcePage[fieldToCopy.FieldName] is Dictionary<string, object>)
                                        {
                                            var taxDictionaryList = (pageTransformationInformation.SourcePage[fieldToCopy.FieldName] as Dictionary<string, object>);
                                            var valueCollectionToCopy = taxDictionaryList["_Child_Items_"] as Object[];

                                            List<string> taxonomyFieldValueArray = new List<string>();
                                            for (int i = 0; i < valueCollectionToCopy.Length; i++)
                                            {
                                                var taxDictionary = valueCollectionToCopy[i] as Dictionary<string, object>;
                                                taxonomyFieldValueArray.Add($"-1;#{taxDictionary["Label"].ToString()}|{taxDictionary["TermGuid"].ToString()}");
                                            }
                                            var valueCollection = new TaxonomyFieldValueCollection(this.sourceClientContext, string.Join(";#", taxonomyFieldValueArray), taxField);
                                            taxField.SetFieldValueByValueCollection(targetPage.PageListItem, valueCollection);
                                        }

                                        isDirty = true;
                                        LogInfo($"{LogStrings.TransformCopyingMetaDataField} {fieldToCopy.FieldName}", LogStrings.Heading_CopyingPageMetadata);
                                    }
                                }
                                break;
                            }
                        case "TaxonomyFieldType":
                            {
                                var taxFieldBeforeCast = pagesLibrary.Fields.Where(p => p.Id.Equals(fieldToCopy.FieldId)).FirstOrDefault();
                                if (taxFieldBeforeCast != null)
                                {
                                    var taxField = this.sourceClientContext.CastTo<TaxonomyField>(taxFieldBeforeCast);
                                    var taxValue = new TaxonomyFieldValue();
                                    if (pageTransformationInformation.SourcePage[fieldToCopy.FieldName] != null)
                                    {
                                        if (pageTransformationInformation.SourcePage[fieldToCopy.FieldName] is TaxonomyFieldValue)
                                        {

                                            taxValue.Label = (pageTransformationInformation.SourcePage[fieldToCopy.FieldName] as TaxonomyFieldValue).Label;
                                            taxValue.TermGuid = (pageTransformationInformation.SourcePage[fieldToCopy.FieldName] as TaxonomyFieldValue).TermGuid;
                                            taxValue.WssId = -1;
                                        }
                                        else if (pageTransformationInformation.SourcePage[fieldToCopy.FieldName] is Dictionary<string, object>)
                                        {
                                            var taxDictionary = (pageTransformationInformation.SourcePage[fieldToCopy.FieldName] as Dictionary<string, object>);
                                            taxValue.Label = taxDictionary["Label"].ToString();
                                            taxValue.TermGuid = taxDictionary["TermGuid"].ToString();
                                            taxValue.WssId = -1;
                                        }
                                        taxField.SetFieldValueByValue(targetPage.PageListItem, taxValue);
                                        isDirty = true;
                                        LogInfo($"{LogStrings.TransformCopyingMetaDataField} {fieldToCopy.FieldName}", LogStrings.Heading_CopyingPageMetadata);
                                    }
                                }
                                break;
                            }
                    }
                }

                if (isDirty)
                {
                    targetPage.PageListItem.Update();
                    this.sourceClientContext.Load(targetPage.PageListItem);
                    this.sourceClientContext.ExecuteQueryRetry();
                }
            }
        }

        /// <summary>
        /// Gets the version of the assembly
        /// </summary>
        /// <returns></returns>
        internal string GetVersion()
        {
            try
            {
                var coreAssembly = Assembly.GetExecutingAssembly();
                return ((AssemblyFileVersionAttribute)coreAssembly.GetCustomAttribute(typeof(AssemblyFileVersionAttribute))).Version.ToString();
            }
            catch (Exception ex)
            {
                LogError(LogStrings.Error_GetVersionError, LogStrings.Heading_GetVersion, ex, true);
            }

            return "undefined";
        }

        internal void InitMeasurement()
        {
            try
            {
                if (System.IO.File.Exists(ExecutionLog))
                {
                    System.IO.File.Delete(ExecutionLog);
                }
            }
            catch { }
        }

        internal void Start()
        {
            watch = Stopwatch.StartNew();
        }

        internal void Stop(string method)
        {
            watch.Stop();
            var elapsedTime = watch.ElapsedMilliseconds;
            System.IO.File.AppendAllText(ExecutionLog, $"{method};{elapsedTime}{Environment.NewLine}");
        }

        /// <summary>
        /// Loads the telemetry and properties for the client object
        /// </summary>
        /// <param name="clientContext"></param>
        internal void LoadClientObject(ClientContext clientContext)
        {
            if (clientContext != null)
            {
                clientContext.ClientTag = $"SPDev:PageTransformator";
                // Load all web properties needed further one
                clientContext.Load(clientContext.Web, p => p.Id, p => p.ServerRelativeUrl, p => p.RootFolder.WelcomePage, p => p.Url, p => p.WebTemplate, p => p.Language);
                clientContext.Load(clientContext.Site, p => p.RootWeb.ServerRelativeUrl, p => p.Id, p => p.Url);
                // Use regular ExecuteQuery as we want to send this custom clienttag
                clientContext.ExecuteQuery();
            }
        }

        /// <summary>
        /// Validates settings when doing a cross farm transformation
        /// </summary>
        /// <param name="baseTransformationInformation">Transformation Information</param>
        /// <remarks>Will disable feature if not supported</remarks>
        internal void CrossFarmTransformationValidation(BaseTransformationInformation baseTransformationInformation)
        {
            // Source only context - allow item level permissions
            // Source to target same base address - allow item level permissions
            // Source to target difference base address - disallow item level permissions

            if(targetClientContext != null && sourceClientContext != null && baseTransformationInformation.KeepPageSpecificPermissions)
            {
                var sourceUrl = sourceClientContext.Url.GetBaseUrl();
                var targetUrl = targetClientContext.Url.GetBaseUrl();

                // Override the setting for keeping item level permissions
                if(!sourceUrl.Equals(targetUrl, StringComparison.InvariantCultureIgnoreCase))
                {
                    baseTransformationInformation.KeepPageSpecificPermissions = false;
                    LogWarning(LogStrings.Warning_ContextValidationFailWithKeepPermissionsEnabled, LogStrings.Heading_InputValidation);

                    // Set a global flag to indicate this is a cross farm transformation (on-prem to SPO tenant or SPO Tenant A to SPO Tenant B)
                    baseTransformationInformation.IsCrossFarmTransformation = true;
                }
            }

            if (sourceClientContext != null)
            {
                baseTransformationInformation.SourceVersion = GetVersion(sourceClientContext);
            }

            if (targetClientContext != null)
            {
                baseTransformationInformation.TargetVersion = GetVersion(targetClientContext);
            }

            if (sourceClientContext != null && targetClientContext == null)
            {
                baseTransformationInformation.TargetVersion = baseTransformationInformation.SourceVersion;
            }

        }
        #endregion

        #region Helper methods
        private SPVersion GetVersion(ClientRuntimeContext clientContext)
        {
            try
            {
                Uri urlUri = new Uri(clientContext.Url);
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create($"{urlUri.Scheme}://{urlUri.DnsSafeHost}:{urlUri.Port}/_vti_pvt/service.cnf");
                request.UseDefaultCredentials = true;

                var response = request.GetResponse();

                using (var dataStream = response.GetResponseStream())
                {
                    // Open the stream using a StreamReader for easy access.
                    using (System.IO.StreamReader reader = new System.IO.StreamReader(dataStream))
                    {
                        // Read the content.Will be in this format
                        // SPO:
                        //vti_encoding: SR | utf8 - nl
                        //vti_extenderversion: SR | 16.0.0.8929
                        // SP2019:
                        // vti_encoding:SR|utf8-nl
                        // vti_extenderversion:SR|16.0.0.10340
                        // SP2016:
                        // vti_encoding: SR | utf8 - nl
                        // vti_extenderversion: SR | 16.0.0.4732
                        // SP2013:
                        // vti_encoding:SR|utf8-nl
                        // vti_extenderversion: SR | 15.0.0.4505
                        // Version numbers from https://buildnumbers.wordpress.com/sharepoint/

                        string version = reader.ReadToEnd().Split('|')[2].Trim();

                        if (Version.TryParse(version, out Version v))
                        {
                            if (v.Major == 14)
                            {
                                return SPVersion.SP2010;
                            }
                            else if (v.Major == 15)
                            {
                                return SPVersion.SP2013;
                            }
                            else if (v.Major == 16)
                            {
                                if (v.MinorRevision < 6000)
                                {
                                    return SPVersion.SP2016;
                                }
                                // As of July 2019 SPO buildnumbers will be higher than 18000 
                                else if (v.MinorRevision > 10300 && v.MinorRevision < 18000)
                                {
                                    return SPVersion.SP2019;
                                }
                                else
                                {
                                    return SPVersion.SPO;
                                }
                            }
                        }
                    }
                }
            }
            catch (WebException ex)
            {
                // todo
            }

            return SPVersion.SPO;
        }
        #endregion
    }
}
