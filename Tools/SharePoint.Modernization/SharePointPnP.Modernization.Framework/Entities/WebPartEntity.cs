using Newtonsoft.Json;
using SharePointPnP.Modernization.Framework.Extensions;
using System;
using System.Collections.Generic;

namespace SharePointPnP.Modernization.Framework.Entities
{
    /// <summary>
    /// Entity to describe a web part on a wiki or webpart page
    /// </summary>
    public class WebPartEntity
    {
        public WebPartEntity()
        {
            this.Properties = new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase);
        }

        public string Type { get; set; }
        public Guid Id { get; set; }
        public string ServerControlId { get; set; }
        public string ZoneId { get; set; }
        public bool Hidden { get; set; }
        public bool IsClosed { get; set; }
        public string Title { get; set; }
        public int Row { get; set; }
        public int Column { get; set; }
        public int Order { get; set; }
        public uint ZoneIndex { get; set; }
        public Dictionary<string, string> Properties { get; set; }

        /// <summary>
        /// Shortened web part type name
        /// </summary>
        /// <returns></returns>
        public string TypeShort()
        {
            return Type.GetTypeShort();
        }

        /// <summary>
        /// Returns this instance as Json
        /// </summary>
        /// <returns>Json serialized string of this web part instance</returns>
        public string Json()
        {
            return JsonConvert.SerializeObject(this);
        }

    }
}
