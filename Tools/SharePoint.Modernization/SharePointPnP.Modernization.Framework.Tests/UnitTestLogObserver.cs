using SharePointPnP.Modernization.Framework.Telemetry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointPnP.Modernization.Framework.Tests
{
    public class UnitTestLogObserver : ILogObserver
    {
        public void Debug(LogEntry entry)
        {
            Console.WriteLine($"DEBUG: HEADING: {entry.Heading}  - Message: {entry.Message} \n\t Source: {entry.Source}, Id: {entry.CorrelationId}");
        }

        public void Error(LogEntry entry)
        {
            var error = entry.Exception != null ? entry.Exception.Message : "No error logged";
            Console.WriteLine($"ERROR: HEADING: {entry.Heading} Message: {entry.Message} \n\t Source: {entry.Source}, Id: {entry.CorrelationId}, Error: { error }");
        }

        public void Info(LogEntry entry)
        {
            Console.WriteLine($"INFO: HEADING: {entry.Heading} Message: {entry.Message} \n\t Source: {entry.Source}, Id: {entry.CorrelationId}");
        }

        public void Warning(LogEntry entry)
        {
            Console.WriteLine($"WARNING: HEADING: {entry.Heading} Message: {entry.Message} \n\t Source: {entry.Source}, Id: {entry.CorrelationId}");
        }
    }
}
