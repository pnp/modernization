using Newtonsoft.Json;
using System;
using System.Collections.Generic;

namespace SharePointPnP.Modernization.Framework.Entities
{
    /// <summary>
    /// Entity to describe a web part on a wiki or webpart page called from web services
    /// </summary>
    public class WebServiceWebPartProperties
    {
        public WebServiceWebPartProperties()
        {
            this.Properties = new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase);
        }

        public string Type { get; set; }
        
        public Guid Id { get; set; }

        public string ControlId { get; set; }
        
        public Dictionary<string, string> Properties { get; set; }

        public string ZoneId { get; set; }

        /// <summary>
        /// Shortened web part type name
        /// </summary>
        /// <returns></returns>
        public string TypeShort()
        {
            string name = Type;
            var typeSplit = Type.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            if (typeSplit.Length > 0)
            {
                name = typeSplit[0];
            }

            return $"{name}";
        }

        public Dictionary<string, object> PropertiesAsStringObjectDictionary()
        {
            Dictionary<string, object> castedCollection = new Dictionary<string, object>();

            foreach (var item in this.Properties)
            {
                castedCollection.Add(item.Key, (object)item.Value);
            }

            return castedCollection;
        }

    }
}
