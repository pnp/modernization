using Microsoft.SharePoint.Client;
using SharePointPnP.Modernization.Framework.Telemetry;
using SharePointPnP.Modernization.Framework.Transform;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Pages;
using SharePointPnP.Modernization.Framework.Cache;

namespace SharePointPnP.Modernization.Framework.Publishing
{
    public class PublishingMetadataTransformator: BaseTransform
    {
        PublishingPageTransformationInformation publishingPageTransformationInformation;
        ClientContext sourceClientContext;
        ClientContext targetClientContext;
        ClientSidePage page;
        PageLayout pageLayoutMappingModel;

        #region Construction
        public PublishingMetadataTransformator(PublishingPageTransformationInformation publishingPageTransformationInformation, ClientContext sourceClientContext, ClientContext targetClientContext, ClientSidePage page, PageLayout publishingPageLayoutModel, IList<ILogObserver> logObservers = null)
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
                foreach (var fieldToProcess in this.pageLayoutMappingModel.MetaData)
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

                        if (targetFieldData.FieldType != "TaxonomyFieldTypeMulti" && targetFieldData.FieldType != "TaxonomyFieldType")
                        {
                            if (this.publishingPageTransformationInformation.SourcePage[fieldToProcess.Name] != null)
                            {
                                this.page.PageListItem[targetFieldData.FieldName] = this.publishingPageTransformationInformation.SourcePage[fieldToProcess.Name];
                                isDirty = true;

                                LogInfo($"{LogStrings.TransformCopyingMetaDataField} {targetFieldData.FieldName}", LogStrings.Heading_CopyingPageMetadata);
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

/*
                // Handle the taxonomy fields
                foreach (var fieldToProcess in this.pageLayoutMappingModel.MetaData)
                {
                    // Process only fields which have a target field set...
                    if (!string.IsNullOrEmpty(fieldToProcess.TargetFieldName))
                    {
                        // Get information about this content type field
                        var targetFieldData = CacheManager.Instance.GetPublishingContentTypeField(this.page.PageListItem.ParentList, contentTypeId, fieldToProcess.TargetFieldName);

                        if (targetFieldData.FieldType == "TaxonomyFieldTypeMulti" || targetFieldData.FieldType == "TaxonomyFieldType")
                        {
                            switch (targetFieldData.FieldType)
                            {
                                case "TaxonomyFieldTypeMulti":
                                    {
                                        var taxFieldBeforeCast = pagesLibrary.Fields.Where(p => p.Id.Equals(targetFieldData.FieldId)).FirstOrDefault();
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
*/
            }
            else
            {
                // TODO: add logging
            }
        }

    }
}
