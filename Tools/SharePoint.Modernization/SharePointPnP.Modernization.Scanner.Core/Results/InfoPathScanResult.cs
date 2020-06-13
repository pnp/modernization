using OfficeDevPnP.Core.Entities;
using System;
using System.Collections.Generic;

namespace SharePoint.Modernization.Scanner.Core.Results
{
    public class InfoPathScanResult: Scan
    {
        public string ListUrl { get; set; }

        public string ListTitle { get; set; }

        public Guid ListId { get; set; }

        /// <summary>
        ///  Indicates how InfoPath is used here: form library or customization of the list form pages
        /// </summary>
        public string InfoPathUsage { get; set; }

        public string InfoPathTemplate { get; set; }

        public bool Enabled { get; set; }

        public int ItemCount { get; set; }

        public DateTime LastItemUserModifiedDate { get; set; }

        /// <summary>
        /// Site administrators
        /// </summary>
        public List<UserEntity> Admins { get; set; }
        /// <summary>
        /// Site owners
        /// </summary>
        public List<UserEntity> Owners { get; set; }
    }
}
