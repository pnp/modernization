using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Pages;
using SharePointPnP.Modernization.Framework.Cache;
using SharePointPnP.Modernization.Framework.Telemetry;
using SharePointPnP.Modernization.Framework.Transform;

namespace SharePointPnP.Modernization.Framework.Publishing
{
    public class PublishingPageHeaderTransformator: BaseTransform
    {
        private PublishingPageTransformationInformation publishingPageTransformationInformation;
        private PublishingPageTransformation publishingPageTransformation;
        private PublishingFunctionProcessor functionProcessor;
        private ClientContext sourceClientContext;
        private ClientContext targetClientContext;

        #region Construction
        public PublishingPageHeaderTransformator(PublishingPageTransformationInformation publishingPageTransformationInformation, ClientContext sourceClientContext, ClientContext targetClientContext, PublishingPageTransformation publishingPageTransformation)
        {
            this.publishingPageTransformationInformation = publishingPageTransformationInformation;
            this.publishingPageTransformation = publishingPageTransformation;
            this.sourceClientContext = sourceClientContext;
            this.targetClientContext = targetClientContext;
            this.functionProcessor = new PublishingFunctionProcessor(publishingPageTransformationInformation.SourcePage, sourceClientContext, targetClientContext, this.publishingPageTransformation);
        }
        #endregion


        #region Header transformation
        public void TransformHeader(ref ClientSidePage targetPage)
        {
            // Get the mapping model to use as it describes how the page header needs to be generated
            string usedPageLayout = System.IO.Path.GetFileNameWithoutExtension(this.publishingPageTransformationInformation.SourcePage.PageLayoutFile());
            var publishingPageTransformationModel = this.publishingPageTransformation.PageLayouts.Where(p => p.Name.Equals(usedPageLayout, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();

            // No layout provided via either the default mapping or custom mapping file provided
            if (publishingPageTransformationModel == null)
            {
                publishingPageTransformationModel = CacheManager.Instance.GetPageLayoutMapping(publishingPageTransformationInformation.SourcePage);
            }

            // Configure the page header
            if (publishingPageTransformationModel.PageHeader == PageLayoutPageHeader.None)
            {
                targetPage.RemovePageHeader();
            }
            else if (publishingPageTransformationModel.PageHeader == PageLayoutPageHeader.Default)
            {
                targetPage.SetDefaultPageHeader();
            }
            else
            {
                // Custom page header

                // ImageServerRelativeUrl 
                string imageServerRelativeUrl = "";
                HeaderField imageServerRelativeUrlField = GetHeaderField(publishingPageTransformationModel, "ImageServerRelativeUrl");

                if (imageServerRelativeUrlField != null)
                {
                    if (!string.IsNullOrEmpty(imageServerRelativeUrlField.Functions))
                    {
                        // execute function
                        var evaluatedField = this.functionProcessor.Process(imageServerRelativeUrlField.Functions, imageServerRelativeUrlField.Name, "string");
                        if (!string.IsNullOrEmpty(evaluatedField.Item1))
                        {
                            imageServerRelativeUrl = evaluatedField.Item2;
                        }
                    }
                    else
                    {
                        imageServerRelativeUrl = this.publishingPageTransformationInformation.SourcePage.FieldValues[imageServerRelativeUrlField.Name]?.ToString().Trim();
                    }
                }

                // Did we get a header image url?
                if (!string.IsNullOrEmpty(imageServerRelativeUrl))
                {
                    string newHeaderImageServerRelativeUrl = "";
                    try
                    {
                        // Integrate asset transformator

                        // Check if the asset lives in the current site...else assume it lives in the rootweb of the site collection
                        ClientContext contextForAssetTransfer = this.sourceClientContext;
                        string assetServerRelativePath = imageServerRelativeUrl.Substring(0, imageServerRelativeUrl.LastIndexOf("/"));
                        string sourceWebRelativePath = this.sourceClientContext.Web.EnsureProperty(p => p.ServerRelativeUrl);

                        if (!assetServerRelativePath.StartsWith(sourceWebRelativePath, StringComparison.InvariantCultureIgnoreCase))
                        {
                            string rootWebUrl = this.sourceClientContext.Site.Url;
                            contextForAssetTransfer = this.sourceClientContext.Clone(rootWebUrl);
                        }

                        // Copy the asset
                        AssetTransfer assetTransfer = new AssetTransfer(contextForAssetTransfer, this.targetClientContext, base.RegisteredLogObservers);
                        newHeaderImageServerRelativeUrl = assetTransfer.TransferAsset(imageServerRelativeUrl, targetPage.PageTitle);
                    }
                    catch (Exception ex)
                    {
                        // TODO: update strings
                        //LogError(LogStrings.Error_ReturnCrossSiteRelativePath, LogStrings.Heading_BuiltInFunctions, ex);
                    }

                    if (!string.IsNullOrEmpty(newHeaderImageServerRelativeUrl))
                    {
                        targetPage.SetCustomPageHeader(newHeaderImageServerRelativeUrl);

                        HeaderField topicHeaderField = GetHeaderField(publishingPageTransformationModel, "TopicHeader");
                        if (topicHeaderField != null)
                        {
                            if (publishingPageTransformationInformation.SourcePage.FieldExistsAndUsed(topicHeaderField.Name))
                            {
                                targetPage.PageHeader.TopicHeader = publishingPageTransformationInformation.SourcePage[topicHeaderField.Name].ToString();
                                targetPage.PageHeader.ShowTopicHeader = true;
                            }
                        }                        
                    }
                    else
                    {
                        // let's fall back to no header
                        targetPage.RemovePageHeader();
                    }
                }
                else
                {
                    // let's fall back to no header
                    targetPage.RemovePageHeader();
                }
            }
        }

        private static HeaderField GetHeaderField(PageLayout publishingPageTransformationModel, string fieldName)
        {
            return publishingPageTransformationModel.Header.Field.Where(p => p.HeaderProperty.Equals(fieldName, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
        }
        #endregion



    }
}
