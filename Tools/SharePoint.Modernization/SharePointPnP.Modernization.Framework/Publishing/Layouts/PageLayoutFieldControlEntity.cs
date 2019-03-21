using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointPnP.Modernization.Framework.Publishing.Layouts
{
    /// <summary>
    /// Class for holding data properties for field control entity to the target web part
    /// </summary>
    public class PageLayoutFieldControlEntity
    {
        public string TargetWebPart { get; set; }
        public string FieldName { get; set; }
        public string Name { get; set; }
        public string FieldType { get; set; }
        public string ProcessFunction { get; set; }
    }
}
