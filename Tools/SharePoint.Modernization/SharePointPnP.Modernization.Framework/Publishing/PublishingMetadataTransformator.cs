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
    public class PublishingMetadataTransformator: BaseTransform
    {
        private PublishingPageTransformationInformation publishingPageTransformationInformation;
        private ClientContext targetClientContext;
        private ClientSidePage page;
        private PageLayout pageLayoutMappingModel;
        private PublishingPageTransformation publishingPageTransformation;
        private PublishingFunctionProcessor functionProcessor;

        #region Construction
        public PublishingMetadataTransformator(PublishingPageTransformationInformation publishingPageTransformationInformation, ClientContext sourceClientContext, ClientContext targetClientContext, ClientSidePage page, PageLayout publishingPageLayoutModel, PublishingPageTransformation publishingPageTransformation, IList<ILogObserver> logObservers = null)
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
            this.targetClientContext = targetClientContext;
            this.page = page;
            this.pageLayoutMappingModel = publishingPageLayoutModel;
            this.publishingPageTransformation = publishingPageTransformation;
            this.functionProcessor = new PublishingFunctionProcessor(publishingPageTransformationInformation.SourcePage, sourceClientContext, targetClientContext, this.publishingPageTransformation, publishingPageTransformationInformation as BaseTransformationInformation, base.RegisteredLogObservers);
        }
        #endregion

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
                        this.page.PageListItem.Update();
                        isDirty = true;
                    }
                }

                // Determine content type to use
                if (string.IsNullOrEmpty(contentTypeId))
                {
                    // grab the default content type
                    contentTypeId = this.page.PageListItem[Constants.ContentTypeIdField].ToString();
                }

                // Copy the field metadata
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
                            if (targetFieldData.FieldType != "TaxonomyFieldTypeMulti" && targetFieldData.FieldType != "TaxonomyFieldType")
                            {
                                if (this.publishingPageTransformationInformation.SourcePage[fieldToProcess.Name] != null)
                                {
                                    if (!string.IsNullOrEmpty(fieldToProcess.Functions))
                                    {
                                        // execute function
                                        var evaluatedField = this.functionProcessor.Process(fieldToProcess.Functions, fieldToProcess.Name, CastToPublishingFunctionProcessorFieldType(targetFieldData.FieldType));
                                        if (!string.IsNullOrEmpty(evaluatedField.Item1))
                                        {
                                            this.page.PageListItem[targetFieldData.FieldName] = evaluatedField.Item2;
                                        }
                                    }
                                    else
                                    {
                                        this.page.PageListItem[targetFieldData.FieldName] = this.publishingPageTransformationInformation.SourcePage[fieldToProcess.Name];
                                    }
                                    isDirty = true;

                                    LogInfo($"{LogStrings.TransformCopyingMetaDataField} {targetFieldData.FieldName}", LogStrings.Heading_CopyingPageMetadata);
                                }
                            }
                        }
                    }
                }

                // Persist changes
                if (isDirty)
                {
                    this.page.PageListItem.Update();
                    targetClientContext.Load(this.page.PageListItem);
                    targetClientContext.ExecuteQueryRetry();
                    isDirty = false;
                }


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

                                                if (this.publishingPageTransformationInformation.SourcePage[fieldToProcess.Name] != null)
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
                                                if (this.publishingPageTransformationInformation.SourcePage[fieldToProcess.Name] != null)
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
                    this.page.PageListItem.Update();
                    targetClientContext.Load(this.page.PageListItem);
                    targetClientContext.ExecuteQueryRetry();
                    isDirty = false;
                }

            }
            else
            {
                // TODO: add logging
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
        #endregion

    }
}
