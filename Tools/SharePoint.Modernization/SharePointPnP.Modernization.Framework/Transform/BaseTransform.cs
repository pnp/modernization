using SharePointPnP.Modernization.Framework.Telemetry;
using System.Diagnostics;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointPnP.Modernization.Framework.Transform
{
    public class BaseTransform
    {
        private IList<ILogObserver> _logObservers;
        private Guid _correlationId;

        /// <summary>
        /// Instantiation of base transform class
        /// </summary>
        public BaseTransform()
        {
            _logObservers = new List<ILogObserver>();
            _correlationId = Guid.NewGuid();
        }
        

        /// <summary>
        /// Registers the observer
        /// </summary>
        /// <param name="observer">The observer.</param>
        public void RegisterObserver(ILogObserver observer)
        {
            if (!_logObservers.Contains(observer))
            {
                _logObservers.Add(observer);
            }
        }

        

        /// <summary>
        /// Notifies the observers of error messages
        /// </summary>
        /// <param name="logEntry">The message.</param>
        public void NotifyLogErrorObservers(LogEntry logEntry)
        {
            StackTrace stackTrace = new StackTrace();
            logEntry.Source = stackTrace.GetFrame(1).ToString();

            foreach (ILogObserver observer in _logObservers)
            {
                observer.Error(logEntry);
            }
        }

        /// <summary>
        /// Notifies the observers of info messages
        /// </summary>
        /// <param name="logEntry">The message.</param>
        public void NotifyLogInfoObservers(LogEntry logEntry)
        {
            foreach (ILogObserver observer in _logObservers)
            {
                observer.Info(logEntry);
            }
        }

        /// <summary>
        /// Notifies the observers of warning messages
        /// </summary>
        /// <param name="logEntry">The message.</param>
        public void NotifyLogWarningObservers(LogEntry logEntry)
        {
            foreach (ILogObserver observer in _logObservers)
            {
                observer.Warning(logEntry);
            }
        }

        /// <summary>
        /// Notifies the observers of debug messages
        /// </summary>
        /// <param name="logEntry">The message.</param>
        public void NotifyLogDebugObservers(LogEntry logEntry)
        {
            foreach (ILogObserver observer in _logObservers)
            {
                observer.Debug(logEntry);
            }
        }
    }
}
