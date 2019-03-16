using Microsoft.SharePoint.Client;
using SharePointPnP.Modernization.Framework.Telemetry;
using SharePointPnP.Modernization.Framework.Transform;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace SharePointPnP.Modernization.Framework.Publishing
{
    public class PageLayoutAnalyser : BaseTransform
    {
        /*
         * Plan
         *  Read a publishing page or read all the publishing page layouts - need to consider both options
         *  Validate that the client context is a publishing site
         *  Determine page layouts and the associated content type
         *  - Using web part manager scan for web part zones and pre-populated web parts
         *  - Detect for field controls - only the metadata behind these can be transformed without an SPFX web part
         *      - Metadata mapping to web part - only some types will be supported
         *  - Using HTML parser deep analysis of the file to map out detected web parts. These are fixed point in the publishing layout.
         *      - This same method could be used to parse HTML fields for inline web parts
         *  - Generate a layout mapping based on analysis
         *  - Validate the Xml prior to output
         *  - Split into molecules of operation for unit testing
         *  - Detect grid system, table or fabric for layout options - consider...
         */

        private ClientContext _context;
        private PublishingPageTransformation _mapping;
        private string _defaultFileName = "PageLayoutMapping.xml";

        /// <summary>
        /// Analyse Page Layouts class constructor
        /// </summary>
        /// <param name="sourceContext">This should be the context of the source web</param>
        /// <param name="logObservers"></param>
        public PageLayoutAnalyser(ClientContext sourceContext, IList<ILogObserver> logObservers = null)
        {
            // Register observers
            if (logObservers != null){
                foreach (var observer in logObservers){
                    base.RegisterObserver(observer);
                }
            }

            _context = sourceContext;

            _mapping = new PublishingPageTransformation();
        }


        /// <summary>
        /// Main entry point into the class to analyse the page layouts
        /// </summary>
        public void Analyse()
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Perform validation to ensure the source site contains page layouts
        /// </summary>
        public void Validate()
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Gets the page layout for analysis
        /// </summary>
        public void GetPageLayout()
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Determine the page layout from a publishing page
        /// </summary>
        public void GetPageLayoutFromPublishingPage()
        {
            //Note: ListItemExtensions class contains this logic - reuse.
            throw new NotImplementedException();
        }

        /// <summary>
        /// Get Metadata mapping from the page layout associated content type
        /// </summary>
        public void GetAssociatedMetadatafromPageLayoutContentType()
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Get web part zones defined in the page layout
        /// </summary>
        public void GetWebPartZones()
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Get fixed web parts defined in the page layout
        /// </summary>
        public void GetFixedWebPartsFromZones()
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Extract the web parts from the page layout HTML outside of web part zones
        /// </summary>
        public void ExtractWebPartsFromPageLayoutHtml()
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Generate the mapping file to output from the analysis
        /// </summary>
        public string GenerateMappingFile()
        {
            try
            {
                XmlSerializer xmlMapping = new XmlSerializer(typeof(PublishingPageTransformation));

                var mappingFileName = _defaultFileName;

                using (StreamWriter sw = new StreamWriter(mappingFileName, false))
                {
                    xmlMapping.Serialize(sw, _mapping);
                }

                var xmlMappingFileLocation = $"{ Environment.CurrentDirectory }\\{ mappingFileName}";
                LogInfo($"{LogStrings.XmlMappingSavedAs}: {xmlMappingFileLocation}");

                return xmlMappingFileLocation;

            }catch(Exception ex)
            {
                var message = string.Format(LogStrings.Error_CannotWriteToXmlFile, ex.Message, ex.StackTrace);
                Console.WriteLine(message);
                LogError(message, LogStrings.Heading_PageLayoutAnalyser, ex);
            }

            return string.Empty;
        }
    }
}
