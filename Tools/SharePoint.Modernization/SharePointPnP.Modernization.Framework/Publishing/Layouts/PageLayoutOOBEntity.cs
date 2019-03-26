using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointPnP.Modernization.Framework.Publishing.Layouts
{
    public class PageLayoutOOBEntity
    {
        /// <summary>
        /// Page Layout Out of the Box Entity Mapping
        /// </summary>
        public PageLayoutOOBEntity()
        {
            IgnoreMapping = false;
        }

        public OOBLayout Layout { get; set; }
        public string Name { get; set; }
        public string PageLayoutTemplate { get; set; }
        public string PageHeader { get; set; }
        public string PageHeaderType { get; set; }
        public bool IgnoreMapping { get; set; }
    }
}
