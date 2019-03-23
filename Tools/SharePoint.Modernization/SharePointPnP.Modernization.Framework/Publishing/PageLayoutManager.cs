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
        /// Loads a page layout mapping file
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

        /// <summary>
        /// Load the default page layout mapping file
        /// </summary>
        /// <returns>A <see cref="PublishingPageTransformation"/> instance.</returns>
        internal PublishingPageTransformation LoadDefaultPageLayoutMappingFile()
        {
            var fileContent = "";
            using (Stream stream = typeof(PageLayoutManager).Assembly.GetManifestResourceStream("SharePointPnP.Modernization.Framework.Publishing.pagelayoutmapping.xml"))
            {
                using (StreamReader reader = new StreamReader(stream))
                {
                    fileContent = reader.ReadToEnd();
                }
            }

            XmlSerializer xmlMapping = new XmlSerializer(typeof(PublishingPageTransformation));
            using (var stream = GenerateStreamFromString(fileContent))
            {
                return (PublishingPageTransformation)xmlMapping.Deserialize(stream);
            }          
        }

        internal PublishingPageTransformation MergePageLayoutMappingFiles(PublishingPageTransformation oobMapping, PublishingPageTransformation customMapping)
        {
            PublishingPageTransformation merged = new PublishingPageTransformation();

            // Handle the page layouts
            List<PageLayout> pageLayouts = new List<PageLayout>();
            foreach (var oobPageLayout in oobMapping.PageLayouts.ToList())
            {
                // If there's the same page layout used in the custom mapping then that one overrides the default
                if (!customMapping.PageLayouts.Where(p=>p.Name.Equals(oobPageLayout.Name, StringComparison.InvariantCultureIgnoreCase)).Any())
                {
                    pageLayouts.Add(oobPageLayout);
                }
            }
            
            // Take over the custom ones
            pageLayouts.AddRange(customMapping.PageLayouts);
            merged.PageLayouts = pageLayouts.ToArray();

            // Handle the add-ons
            merged.AddOns = customMapping.AddOns;

            return merged;
        }

        #region Helper methods
        /// <summary>
        /// Transforms a string into a stream
        /// </summary>
        /// <param name="s">String to transform</param>
        /// <returns>Stream</returns>
        private static Stream GenerateStreamFromString(string s)
        {
            var stream = new MemoryStream();
            var writer = new StreamWriter(stream);
            writer.Write(s);
            writer.Flush();
            stream.Position = 0;
            return stream;
        }
        #endregion
    }
}
