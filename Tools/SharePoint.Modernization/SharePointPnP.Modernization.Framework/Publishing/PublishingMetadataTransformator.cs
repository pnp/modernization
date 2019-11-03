using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using OfficeDevPnP.Core.Pages;
using OfficeDevPnP.Core.Utilities;
using SharePointPnP.Modernization.Framework.Cache;
using SharePointPnP.Modernization.Framework.Telemetry;
using SharePointPnP.Modernization.Framework.Transform;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SharePointPnP.Modernization.Framework.Publishing
{
    /// <summary>
    /// Class responsible for handling the metadata copying of publishing pages
    /// </summary>
    public class PublishingMetadataTransformator : BaseTransform
    {
        private PublishingPageTransformationInformation publishingPageTransformationInformation;
        private ClientContext sourceClientContext;
        private ClientContext targetClientContext;
        private ClientSidePage page;
        private PageLayout pageLayoutMappingModel;
        private PublishingPageTransformation publishingPageTransformation;
        private PublishingFunctionProcessor functionProcessor;
        private UserTransformator userTransformator;
        
        #region Construction
        public PublishingMetadataTransformator(PublishingPageTransformationInformation publishingPageTransformationInformation, ClientContext sourceClientContext, 
            ClientContext targetClientContext, ClientSidePage page, PageLayout publishingPageLayoutModel, PublishingPageTransformation publishingPageTransformation,
            UserTransformator userTransformator,IList<ILogObserver> logObservers = null)
        {
            // Register observers
            if (logObservers != null)
            {
                foreach (var observer in logObservers)
                {
                    base.RegisterObserver(observer);
                }
            }

            this.publishingPageTransformationInformation = publishingPageTransformationInformation;
            this.sourceClientContext = sourceClientContext;
            this.targetClientContext = targetClientContext;
            this.page = page;
            this.pageLayoutMappingModel = publishingPageLayoutModel;
            this.publishingPageTransformation = publishingPageTransformation;
            this.functionProcessor = new PublishingFunctionProcessor(publishingPageTransformationInformation.SourcePage, sourceClientContext, targetClientContext, this.publishingPageTransformation, publishingPageTransformationInformation as BaseTransformationInformation, base.RegisteredLogObservers);
            this.userTransformator = userTransformator;
        }
        #endregion

        /// <summary>
        /// Process the metadata copying (as defined in the used page layout mapping)
        /// </summary>
        public void Transform()
        {
            if (this.pageLayoutMappingModel != null)
            {
                bool isDirty = false;
                bool listItemWasReloaded = false;
                string contentTypeId = null;

                // Set content type
                if (!string.IsNullOrEmpty(this.pageLayoutMappingModel.AssociatedContentType))
                {
                    contentTypeId = CacheManager.Instance.GetContentTypeId(this.page.PageListItem.ParentList, pageLayoutMappingModel.AssociatedContentType);
                    if (!string.IsNullOrEmpty(contentTypeId))
                    {
                        // Load the target page list item, needs to be loaded as it was previously saved and we need to avoid version conflicts
                        this.targetClientContext.Load(this.page.PageListItem);
                        this.targetClientContext.ExecuteQueryRetry();
                        listItemWasReloaded = true;

                        this.page.PageListItem[Constants.ContentTypeIdField] = contentTypeId;
                        this.page.PageListItem.UpdateOverwriteVersion();
                        isDirty = true;
                    }
                }

                // Determine content type to use
                if (string.IsNullOrEmpty(contentTypeId))
                {
                    // grab the default content type
                    contentTypeId = this.page.PageListItem[Constants.ContentTypeIdField].ToString();
                }

                if (this.pageLayoutMappingModel.MetaData != null)
                {
                    // Handle the taxonomy fields
                    bool targetSitePagesLibraryLoaded = false;
                    List targetSitePagesLibrary = null;
                    foreach (var fieldToProcess in this.pageLayoutMappingModel.MetaData.Field)
                    {
                        // Process only fields which have a target field set...
                        if (!string.IsNullOrEmpty(fieldToProcess.TargetFieldName))
                        {
                            if (!listItemWasReloaded)
                            {
                                // Load the target page list item, needs to be loaded as it was previously saved and we need to avoid version conflicts
                                this.targetClientContext.Load(this.page.PageListItem);
                                this.targetClientContext.ExecuteQueryRetry();
                                listItemWasReloaded = true;
                            }

                            // Get information about this content type field
                            var targetFieldData = CacheManager.Instance.GetPublishingContentTypeField(this.page.PageListItem.ParentList, contentTypeId, fieldToProcess.TargetFieldName);

                            if (targetFieldData == null)
                            {
                                LogWarning($"{LogStrings.TransformCopyingMetaDataFieldSkipped} {fieldToProcess.TargetFieldName}", LogStrings.Heading_CopyingPageMetadata);
                            }
                            else
                            {
                                if (targetFieldData.FieldType == "TaxonomyFieldTypeMulti" || targetFieldData.FieldType == "TaxonomyFieldType")
                                {
                                    if (!targetSitePagesLibraryLoaded)
                                    {
                                        var sitePagesServerRelativeUrl = UrlUtility.Combine(targetClientContext.Web.ServerRelativeUrl, "sitepages");
                                        targetSitePagesLibrary = this.targetClientContext.Web.GetList(sitePagesServerRelativeUrl);
                                        this.targetClientContext.Web.Context.Load(targetSitePagesLibrary, l => l.Fields.IncludeWithDefaultProperties(f => f.Id, f => f.Title, f => f.Hidden, f => f.InternalName, f => f.DefaultValue, f => f.Required));
                                        this.targetClientContext.ExecuteQueryRetry();
                                        targetSitePagesLibraryLoaded = true;
                                    }

                                    switch (targetFieldData.FieldType)
                                    {
                                        case "TaxonomyFieldTypeMulti":
                                            {
                                                var taxFieldBeforeCast = targetSitePagesLibrary.Fields.Where(p => p.Id.Equals(targetFieldData.FieldId)).FirstOrDefault();
                                                if (taxFieldBeforeCast != null)
                                                {
                                                    var taxField = this.targetClientContext.CastTo<TaxonomyField>(taxFieldBeforeCast);

                                                    if (this.publishingPageTransformationInformation.SourcePage.FieldExists(fieldToProcess.Name))
                                                    {
                                                        if (this.publishingPageTransformationInformation.SourcePage[fieldToProcess.Name] is TaxonomyFieldValueCollection)
                                                        {
                                                            var valueCollectionToCopy = (this.publishingPageTransformationInformation.SourcePage[fieldToProcess.Name] as TaxonomyFieldValueCollection);
                                                            var taxonomyFieldValueArray = valueCollectionToCopy.Select(taxonomyFieldValue => $"-1;#{taxonomyFieldValue.Label}|{taxonomyFieldValue.TermGuid}");
                                                            var valueCollection = new TaxonomyFieldValueCollection(this.targetClientContext, string.Join(";#", taxonomyFieldValueArray), taxField);
                                                            taxField.SetFieldValueByValueCollection(this.page.PageListItem, valueCollection);
                                                        }
                                                        else if (this.publishingPageTransformationInformation.SourcePage[fieldToProcess.Name] is Dictionary<string, object>)
                                                        {
                                                            var taxDictionaryList = (this.publishingPageTransformationInformation.SourcePage[fieldToProcess.Name] as Dictionary<string, object>);
                                                            var valueCollectionToCopy = taxDictionaryList["_Child_Items_"] as Object[];

                                                            List<string> taxonomyFieldValueArray = new List<string>();
                                                            for (int i = 0; i < valueCollectionToCopy.Length; i++)
                                                            {
                                                                var taxDictionary = valueCollectionToCopy[i] as Dictionary<string, object>;
                                                                taxonomyFieldValueArray.Add($"-1;#{taxDictionary["Label"].ToString()}|{taxDictionary["TermGuid"].ToString()}");
                                                            }
                                                            var valueCollection = new TaxonomyFieldValueCollection(this.targetClientContext, string.Join(";#", taxonomyFieldValueArray), taxField);
                                                            taxField.SetFieldValueByValueCollection(this.page.PageListItem, valueCollection);
                                                        }

                                                        isDirty = true;
                                                        LogInfo($"{LogStrings.TransformCopyingMetaDataField} {targetFieldData.FieldName}", LogStrings.Heading_CopyingPageMetadata);
                                                    }
                                                    else
                                                    {
                                                        // Log that field in page layout mapping was not found
                                                        LogWarning(string.Format(LogStrings.Warning_FieldNotFoundInSourcePage, fieldToProcess.Name), LogStrings.Heading_CopyingPageMetadata);
                                                    }
                                                }
                                                break;
                                            }
                                        case "TaxonomyFieldType":
                                            {
                                                var taxFieldBeforeCast = targetSitePagesLibrary.Fields.Where(p => p.Id.Equals(targetFieldData.FieldId)).FirstOrDefault();
                                                if (taxFieldBeforeCast != null)
                                                {
                                                    var taxField = this.targetClientContext.CastTo<TaxonomyField>(taxFieldBeforeCast);
                                                    var taxValue = new TaxonomyFieldValue();

                                                    if (this.publishingPageTransformationInformation.SourcePage.FieldExists(fieldToProcess.Name))
                                                    {
                                                        if (this.publishingPageTransformationInformation.SourcePage[fieldToProcess.Name] is TaxonomyFieldValue)
                                                        {

                                                            taxValue.Label = (this.publishingPageTransformationInformation.SourcePage[fieldToProcess.Name] as TaxonomyFieldValue).Label;
                                                            taxValue.TermGuid = (this.publishingPageTransformationInformation.SourcePage[fieldToProcess.Name] as TaxonomyFieldValue).TermGuid;
                                                            taxValue.WssId = -1;
                                                        }
                                                        else if (this.publishingPageTransformationInformation.SourcePage[fieldToProcess.Name] is Dictionary<string, object>)
                                                        {
                                                            var taxDictionary = (this.publishingPageTransformationInformation.SourcePage[fieldToProcess.Name] as Dictionary<string, object>);
                                                            taxValue.Label = taxDictionary["Label"].ToString();
                                                            taxValue.TermGuid = taxDictionary["TermGuid"].ToString();
                                                            taxValue.WssId = -1;
                                                        }
                                                        taxField.SetFieldValueByValue(this.page.PageListItem, taxValue);
                                                        isDirty = true;
                                                        LogInfo($"{LogStrings.TransformCopyingMetaDataField} {targetFieldData.FieldName}", LogStrings.Heading_CopyingPageMetadata);
                                                    }
                                                    else
                                                    {
                                                        // Log that field in page layout mapping was not found
                                                        LogWarning(string.Format(LogStrings.Warning_FieldNotFoundInSourcePage, fieldToProcess.Name), LogStrings.Heading_CopyingPageMetadata);
                                                    }
                                                }
                                                break;
                                            }
                                    }
                                }
                            }
                        }
                    }

                    // Persist changes
                    if (isDirty)
                    {
                        this.page.PageListItem.UpdateOverwriteVersion();
                        targetClientContext.Load(this.page.PageListItem);
                        targetClientContext.ExecuteQueryRetry();
                        isDirty = false;
                    }

                    string bannerImageUrl = null;

                    // Copy the field metadata
                    foreach (var fieldToProcess in this.pageLayoutMappingModel.MetaData.Field)
                    {

                        // check if the source field name attribute contains a delimiter value
                        if (fieldToProcess.Name.Contains(";"))
                        {
                            // extract the array of field names to process, and trims each one
                            string[] sourceFieldNames = fieldToProcess.Name.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray();

                            // sets the field name to the first "valid" entry
                            fieldToProcess.Name = this.publishingPageTransformationInformation.GetFirstNonEmptyFieldName(sourceFieldNames);
                        }

                        // Process only fields which have a target field set...
                        if (!string.IsNullOrEmpty(fieldToProcess.TargetFieldName))
                        {
                            if (!listItemWasReloaded)
                            {
                                // Load the target page list item, needs to be loaded as it was previously saved and we need to avoid version conflicts
                                this.targetClientContext.Load(this.page.PageListItem);
                                this.targetClientContext.ExecuteQueryRetry();
                                listItemWasReloaded = true;
                            }

                            // Get information about this content type field
                            var targetFieldData = CacheManager.Instance.GetPublishingContentTypeField(this.page.PageListItem.ParentList, contentTypeId, fieldToProcess.TargetFieldName);

                            if (targetFieldData == null)
                            {
                                LogWarning($"{LogStrings.TransformCopyingMetaDataFieldSkipped} {fieldToProcess.TargetFieldName}", LogStrings.Heading_CopyingPageMetadata);
                            }
                            else
                            {
                                if (targetFieldData.FieldType != "TaxonomyFieldTypeMulti" && targetFieldData.FieldType != "TaxonomyFieldType")
                                {
                                    if (this.publishingPageTransformationInformation.SourcePage.FieldExists(fieldToProcess.Name))
                                    {
                                        object fieldValueToSet = null;

                                        if (!string.IsNullOrEmpty(fieldToProcess.Functions))
                                        {
                                            // execute function
                                            var evaluatedField = this.functionProcessor.Process(fieldToProcess.Functions, fieldToProcess.Name, CastToPublishingFunctionProcessorFieldType(targetFieldData.FieldType));
                                            if (!string.IsNullOrEmpty(evaluatedField.Item1))
                                            {
                                                fieldValueToSet = evaluatedField.Item2;
                                            }
                                        }
                                        else
                                        {
                                            fieldValueToSet = this.publishingPageTransformationInformation.SourcePage[fieldToProcess.Name];
                                        }

                                        if (fieldValueToSet != null)
                                        {
                                            if (targetFieldData.FieldType == "User" || targetFieldData.FieldType == "UserMulti")
                                            {
                                                if (fieldValueToSet is FieldUserValue)
                                                {
                                                    // Publishing page transformation always goes cross site collection, so we'll need to lookup a user again
                                                    // Important to use a cloned context to not mess up with the pending list item updates
                                                    try
                                                    {
                                                        // Source User
                                                        var fieldUser = (fieldValueToSet as FieldUserValue).LookupValue;
                                                        // Mapped target user
                                                        fieldUser = this.userTransformator.RemapPrincipal(this.sourceClientContext, (fieldValueToSet as FieldUserValue));

                                                        // Ensure user exists on target site
                                                        var ensuredUserOnTarget = CacheManager.Instance.GetEnsuredUser(this.page.Context, fieldUser);
                                                        if (ensuredUserOnTarget != null)
                                                        {
                                                            // Prep a new FieldUserValue object instance and update the list item
                                                            var newUser = new FieldUserValue()
                                                            {
                                                                LookupId = ensuredUserOnTarget.Id
                                                            };
                                                            this.page.PageListItem[targetFieldData.FieldName] = newUser;
                                                        }
                                                        else
                                                        {
                                                            // Clear target field - needed in overwrite scenarios
                                                            this.page.PageListItem[targetFieldData.FieldName] = null;
                                                            LogWarning(string.Format(LogStrings.Warning_UserIsNotMappedOrResolving, (fieldValueToSet as FieldUserValue).LookupValue, targetFieldData.FieldName), LogStrings.Heading_CopyingPageMetadata);
                                                        }
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        LogWarning(string.Format(LogStrings.Warning_UserIsNotResolving, (fieldValueToSet as FieldUserValue).LookupValue, ex.Message), LogStrings.Heading_CopyingPageMetadata);
                                                    }
                                                }
                                                else
                                                {
                                                    List<FieldUserValue> userValues = new List<FieldUserValue>();
                                                    foreach (var currentUser in (fieldValueToSet as Array))
                                                    {
                                                        try
                                                        {
                                                            // Source User
                                                            var fieldUser = (currentUser as FieldUserValue).LookupValue;
                                                            // Mapped target user
                                                            fieldUser = this.userTransformator.RemapPrincipal(this.sourceClientContext, (currentUser as FieldUserValue));

                                                            // Ensure user exists on target site
                                                            var ensuredUserOnTarget = CacheManager.Instance.GetEnsuredUser(this.page.Context, fieldUser);
                                                            if (ensuredUserOnTarget != null)
                                                            {
                                                                // Prep a new FieldUserValue object instance
                                                                var newUser = new FieldUserValue()
                                                                {
                                                                    LookupId = ensuredUserOnTarget.Id
                                                                };

                                                                userValues.Add(newUser);
                                                            }
                                                            else
                                                            {
                                                                LogWarning(string.Format(LogStrings.Warning_UserIsNotMappedOrResolving, (currentUser as FieldUserValue).LookupValue, targetFieldData.FieldName), LogStrings.Heading_CopyingPageMetadata);
                                                            }
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            LogWarning(string.Format(LogStrings.Warning_UserIsNotResolving, (currentUser as FieldUserValue).LookupValue, ex.Message), LogStrings.Heading_CopyingPageMetadata);
                                                        }
                                                    }

                                                    if (userValues.Count > 0)
                                                    {
                                                        this.page.PageListItem[targetFieldData.FieldName] = userValues.ToArray();
                                                    }
                                                    else
                                                    {
                                                        // Clear target field - needed in overwrite scenarios
                                                        this.page.PageListItem[targetFieldData.FieldName] = null;
                                                    }

                                                }                                                
                                            }
                                            else
                                            {
                                                this.page.PageListItem[targetFieldData.FieldName] = fieldValueToSet;

                                                // If we set the BannerImageUrl we also need to update the page to ensure this updated page image "sticks"
                                                if (targetFieldData.FieldName == Constants.BannerImageUrlField)
                                                {
                                                    bannerImageUrl = fieldValueToSet.ToString();
                                                }
                                            }

                                            isDirty = true;

                                            LogInfo($"{LogStrings.TransformCopyingMetaDataField} {targetFieldData.FieldName}", LogStrings.Heading_CopyingPageMetadata);
                                        }
                                    }
                                    else
                                    {
                                        // Log that field in page layout mapping was not found
                                        LogWarning(string.Format(LogStrings.Warning_FieldNotFoundInSourcePage, fieldToProcess.Name), LogStrings.Heading_CopyingPageMetadata);
                                    }
                                }
                            }
                        }
                    }

                    // Persist changes
                    if (isDirty)
                    {
                        // If we've set a custom thumbnail value then we need to update the page html to mark the isDefaultThumbnail pageslicer property to false
                        if (!string.IsNullOrEmpty(bannerImageUrl))
                        {
                            this.page.PageListItem[Constants.CanvasContentField] = SetIsDefaultThumbnail(this.page.PageListItem[Constants.CanvasContentField].ToString());
                        }

                        this.page.PageListItem.UpdateOverwriteVersion();
                        targetClientContext.Load(this.page.PageListItem);
                        targetClientContext.ExecuteQueryRetry();


                        isDirty = false;
                    }
                }
            }
            else
            {
                LogDebug("Page Layout mapping model not found", LogStrings.Heading_CopyingPageMetadata);
            }
        }

        #region Helper methods
        private PublishingFunctionProcessor.FieldType CastToPublishingFunctionProcessorFieldType(string fieldType)
        {
            if (fieldType.Equals("User", StringComparison.InvariantCultureIgnoreCase))
            {
                return PublishingFunctionProcessor.FieldType.User;
            }
            else
            {
                return PublishingFunctionProcessor.FieldType.String;
            }
        }

        private string SetIsDefaultThumbnail(string pageHtml)
        {
            return pageHtml.Replace("&quot;isDefaultThumbnail&quot;&#58;true", "&quot;isDefaultThumbnail&quot;&#58;false");
        }
        #endregion

    }
}
