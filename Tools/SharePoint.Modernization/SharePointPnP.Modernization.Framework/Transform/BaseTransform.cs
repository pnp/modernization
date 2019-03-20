using SharePointPnP.Modernization.Framework.Telemetry;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace SharePointPnP.Modernization.Framework.Transform
{
    /// <summary>
    /// Base logging implementation
    /// </summary>
    public class BaseTransform
    {
        private IList<ILogObserver> _logObservers;
        private Guid _correlationId;

        /// <summary>
        /// List of registered log observers
        /// </summary>
        public IList<ILogObserver> RegisteredLogObservers {
            get{
                return _logObservers;
            }
        }

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
        /// Flush all log observers
        /// </summary>
        public void FlushObservers()
        {
            foreach (ILogObserver observer in _logObservers)
            {
                observer.Flush();
            }
        }

        /// <summary>
        /// Flush Specific Observer of a type
        /// </summary>
        /// <typeparam name="T"></typeparam>
        public void FlushSpecificObserver<T>()
        {
            var observerType = typeof(T);

            foreach (ILogObserver observer in _logObservers)
            {
                if (observer.GetType() == observerType)
                {
                    observer.Flush();
                }
            }
        }
        

        /// <summary>
        /// Notifies the observers of error messages
        /// </summary>
        /// <param name="logEntry">The message.</param>
        public void LogError(string message, string heading = "", Exception exception = null, bool ignoreException = false)
        {
            StackTrace stackTrace = new StackTrace();
            var logEntry = new LogEntry() {
                Heading = heading,
                Message = message,
                CorrelationId = _correlationId,
                Source = stackTrace.GetFrame(1).GetMethod().ToString(),
                Exception = exception,
                IgnoreException = ignoreException
            };

            Log(logEntry, LogLevel.Error);
        }

        /// <summary>
        /// Notifies the observers of info messages
        /// </summary>
        /// <param name="logEntry">The message.</param>
        public void LogInfo(string message, string heading = "")
        {
            StackTrace stackTrace = new StackTrace();
            var logEntry = new LogEntry() { Heading = heading, Message = message, CorrelationId = _correlationId, Source = stackTrace.GetFrame(1).GetMethod().ToString() };

            Log(logEntry, LogLevel.Information);
        }

        /// <summary>
        /// Notifies the observers of warning messages
        /// </summary>
        /// <param name="logEntry">The message.</param>
        public void LogWarning(string message, string heading = "")
        {
            StackTrace stackTrace = new StackTrace();
            var logEntry = new LogEntry() { Heading = heading, Message = message, CorrelationId = _correlationId, Source = stackTrace.GetFrame(1).GetMethod().ToString() };

            Log(logEntry, LogLevel.Warning);
        }

        /// <summary>
        /// Notifies the observers of debug messages
        /// </summary>
        /// <param name="logEntry">The message.</param>
        public void LogDebug(string message, string heading = "")
        {
            StackTrace stackTrace = new StackTrace();
            var logEntry = new LogEntry() { Heading = heading, Message = message, CorrelationId = _correlationId, Source = stackTrace.GetFrame(1).GetMethod().ToString() };

            Log(logEntry, LogLevel.Debug);
        }

        /// <summary>
        /// Log entries into the observers
        /// </summary>
        /// <param name="entry"></param>
        public void Log(LogEntry entry, LogLevel level)
        {
            foreach (ILogObserver observer in _logObservers)
            {
                switch (level)
                {
                    case LogLevel.Debug:
                        observer.Debug(entry);
                        break;
                    case LogLevel.Error:
                        observer.Error(entry);
                        break;
                    case LogLevel.Warning:
                        observer.Warning(entry);
                        break;
                    case LogLevel.Information:
                        observer.Info(entry);
                        break;
                    default:
                        observer.Info(entry);
                        break;
                }
                
            }
        }

        /// <summary>
        /// Sets the page name of the page being transformed
        /// </summary>
        /// <param name="pageName">Name of the page being transformed</param>
        public void SetPage(string pageName)
        {
            foreach (ILogObserver observer in _logObservers)
            {
                observer.SetPage(pageName);
            }
        }
    }
}
