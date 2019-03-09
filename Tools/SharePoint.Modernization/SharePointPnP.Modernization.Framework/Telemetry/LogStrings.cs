using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointPnP.Modernization.Framework.Telemetry
{
    public static class LogStrings
    {
        // Ensure the string variables are meaningful and readable from a developer point of view.
        // Eventually this needs to localised to create multi-lingual outputs.

        public const string KeyValueSeperatorToken = ";#;";

        // Page Transformation
        public const string Heading_PageTransformationInfomation = "Page Transformation Information";
        public const string Heading_Summary = "Summary";

        // Content Transformator
        public const string Heading_ContentTransform = "Content Transform";
        public const string Heading_MappingWebParts = "Web Part Mapping";
        public const string Heading_AddingWebPartsToPage = "Adding Web Parts to Target Page";

    }
}
