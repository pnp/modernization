using SharePointPnP.Modernization.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace SharePoint.Modernization.Scanner.Utilities
{
    /// <summary>
    /// This class handles the interactions with the page transformation model
    /// </summary>
    public class PageTransformationManager
    {

        /// <summary>
        /// Load the page transformation model that will be used
        /// </summary>
        /// <returns></returns>
        public PageTransformation LoadPageTransformationModel()
        {
            PageTransformation pt = null;

            // Load xml mapping data
            XmlSerializer xmlMapping = new XmlSerializer(typeof(PageTransformation));

            // If there's a webpartmapping file in the .exe folder then use that
            if (System.IO.File.Exists("webpartmapping.xml"))
            {
                using (var stream = new FileStream("webpartmapping.xml", FileMode.Open))
                {
                    pt = (PageTransformation)xmlMapping.Deserialize(stream);
                }
            }
            else
            {
                // No webpartmapping file found, let's grab the embedded one
                string webpartMappingString = WebpartMappingLoader.LoadFile("SharePoint.Modernization.Scanner.webpartmapping.xml");
                using (var stream = WebpartMappingLoader.GenerateStreamFromString(webpartMappingString))
                {
                    pt = (PageTransformation)xmlMapping.Deserialize(stream);

                    // Drop web parts that have community mappings as we're not sure if that mapping will be used. This is 
                    // needed to align with the older model where the community mapping was in comments in the standard
                    // mapping file. Providing the webpartmapping file as file in the same folder as the scanner will 
                    // allow to "count" these web parts as transformable (in case a customer wants that)
                    var webPartsWithMappingsToRemove = pt.WebParts.Where(p => p.Type.Equals(WebParts.ScriptEditor) || p.Type.Equals(WebParts.SimpleForm));
                    if (webPartsWithMappingsToRemove.Any())
                    {
                        foreach (var webPart in webPartsWithMappingsToRemove)
                        {
                            webPart.Mappings = null;
                        }
                    }
                }
            }

            return pt;
        }





    }
}
