using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace SharePointPnP.Modernization.Framework.Publishing
{
    /// <summary>
    /// Class used to manage SharePoint Publishing page layouts
    /// </summary>
    public class PageLayoutManager
    {
        private ClientContext sourceContext;
        private ClientContext targetContext;

        #region Construction
        /// <summary>
        /// Constructs the page layout manager class
        /// </summary>
        /// <param name="source">Client context of the source web</param>
        public PageLayoutManager(ClientContext source): this (source, null)
        {
        }

        /// <summary>
        /// Constructs the page layout manager class
        /// </summary>
        /// <param name="source">Client context of the source web</param>
        /// <param name="target">Client context for the target web</param>
        public PageLayoutManager(ClientContext source, ClientContext target)
        {
            this.sourceContext = source ?? throw new ArgumentNullException("Please provide a value for parameter source.");
            // target and source will be set the same in case no target was specified
            this.targetContext = target ?? source;            
        }
        #endregion

        /// <summary>
        /// Loads a page mapping file
        /// </summary>
        /// <param name="pageLayoutMappingFile">Path and name of the page mapping file</param>
        /// <returns>A <see cref="PublishingPageTransformation"/> instance.</returns>
        public PublishingPageTransformation LoadPageLayoutMappingFile(string pageLayoutMappingFile)
        {
            if (!System.IO.File.Exists(pageLayoutMappingFile))
            {
                throw new ArgumentException($"File {pageLayoutMappingFile} does not exist.");
            }

            XmlSerializer xmlMapping = new XmlSerializer(typeof(PublishingPageTransformation));
            using (var stream = new FileStream(pageLayoutMappingFile, FileMode.Open))
            {
                return (PublishingPageTransformation)xmlMapping.Deserialize(stream);
            }
        }

    }
}
