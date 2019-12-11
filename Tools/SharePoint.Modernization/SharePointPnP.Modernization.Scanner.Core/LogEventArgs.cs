using System;

namespace SharePointPnP.Modernization.Scanner.Core
{
    public enum LogSeverity
    {
        Information = 0,
        Warning = 1,
        Error = 2
    }

    public class LogEventArgs: EventArgs
    {
        public string Message { get; set; }
        public LogSeverity Severity { get; set; }
        public DateTime TriggeredAt { get; set; }
    }
}
