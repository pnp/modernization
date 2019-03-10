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

        #region Construction
        public PageLayoutManager(ClientContext context)
        {
            this.sourceContext = context;
        }
        #endregion

        public PublishingPageTransformation ReadPageLayoutMappingFile(string pageLayoutMappingFile)
        {
            XmlSerializer xmlMapping = new XmlSerializer(typeof(PublishingPageTransformation));
            using (var stream = new FileStream(pageLayoutMappingFile, FileMode.Open))
            {
                return (PublishingPageTransformation)xmlMapping.Deserialize(stream);
            }
        }

    }
}
