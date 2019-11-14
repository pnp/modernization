using System.IO;

namespace SharePointPnP.Modernization.Scanner.Core.Reports
{
    public class ReportStream
    {
        public string Source { get; set; }
        public string Name { get; set; }
        public Stream DataStream { get; set; }
    }
}
