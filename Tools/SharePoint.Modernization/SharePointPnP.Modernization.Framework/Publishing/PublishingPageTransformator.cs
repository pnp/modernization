using Microsoft.SharePoint.Client;
using SharePointPnP.Modernization.Framework.Telemetry;
using SharePointPnP.Modernization.Framework.Transform;
using System;
using System.IO;
using System.Xml.Serialization;

namespace SharePointPnP.Modernization.Framework.Publishing
{
    /// <summary>
    /// Transforms a classic publishing page into a modern client side page
    /// </summary>
    public class PublishingPageTransformator : BasePageTransformator
    {
        private PublishingPageTransformation publishingPageTransformation;

        #region Construction
        public PublishingPageTransformator(ClientContext sourceClientContext, ClientContext targetClientContext) : this(sourceClientContext, targetClientContext, "webpartmapping.xml", null)
        {
        }

        public PublishingPageTransformator(ClientContext sourceClientContext, ClientContext targetClientContext, string publishingPageTransformationFile) : this(sourceClientContext, targetClientContext, "webpartmapping.xml", publishingPageTransformationFile)
        {
        }

        public PublishingPageTransformator(ClientContext sourceClientContext, ClientContext targetClientContext, string pageTransformationFile, string publishingPageTransformationFile)
        {
#if DEBUG && MEASURE
            InitMeasurement();
#endif

            this.sourceClientContext = sourceClientContext ?? throw new ArgumentException("sourceClientContext must be provided.");
            this.targetClientContext = targetClientContext ?? throw new ArgumentException("targetClientContext must be provided."); ;

            this.version = GetVersion();
            this.pageTelemetry = new PageTelemetry(version);

            // Load xml mapping data
            XmlSerializer xmlMapping = new XmlSerializer(typeof(PageTransformation));
            using (var stream = new FileStream(pageTransformationFile, FileMode.Open))
            {
                this.pageTransformation = (PageTransformation)xmlMapping.Deserialize(stream);
            }

            // Load the page layout mapping data
            this.publishingPageTransformation = new PageLayoutManager(this.sourceClientContext).LoadPageLayoutMappingFile(publishingPageTransformationFile);
        }

        public PublishingPageTransformator(ClientContext sourceClientContext, ClientContext targetClientContext, PageTransformation pageTransformationModel, PublishingPageTransformation publishingPageTransformationModel)
        {
#if DEBUG && MEASURE
            InitMeasurement();
#endif

            this.sourceClientContext = sourceClientContext ?? throw new ArgumentException("sourceClientContext must be provided.");
            this.targetClientContext = targetClientContext ?? throw new ArgumentException("targetClientContext must be provided."); ;

            this.version = GetVersion();
            this.pageTelemetry = new PageTelemetry(version);

            this.pageTransformation = pageTransformationModel;
            this.publishingPageTransformation = publishingPageTransformationModel;
        }
        #endregion


        /// <summary>
        /// Transform the publishing page
        /// </summary>
        /// <param name="publishingPageTransformationInformation">Information about the publishing page to transform</param>
        /// <returns>The path to the created modern page</returns>
        public string Transform(PublishingPageTransformationInformation publishingPageTransformationInformation)
        {
            // todo

            return "";
        }

    }
}
