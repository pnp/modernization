using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.Extensibility;
using Microsoft.Azure.WebJobs.Host;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace SharePointPnP.ModernizationFunction.Telemetry
{
    public class FunctionTelemetry
    {
        private readonly TelemetryClient telemetryClient;

        #region Construction
        /// <summary>
        /// Instantiates the telemetry client
        /// </summary>
        public FunctionTelemetry(TraceWriter log)
        {
            try
            {
                this.telemetryClient = new TelemetryClient
                {
                    InstrumentationKey = "373400f5-a9cc-48f3-8298-3fd7f4c063d6"
                };

                // Setting this is needed to make metric tracking work
                TelemetryConfiguration.Active.InstrumentationKey = this.telemetryClient.InstrumentationKey;

                this.telemetryClient.Context.Session.Id = Guid.NewGuid().ToString();
                this.telemetryClient.Context.Cloud.RoleInstance = "SharePointPnPPageTransformation";
                this.telemetryClient.Context.Device.OperatingSystem = Environment.OSVersion.ToString();
                var coreAssembly = Assembly.GetExecutingAssembly();
                this.telemetryClient.Context.GlobalProperties.Add("Version", ((AssemblyFileVersionAttribute)coreAssembly.GetCustomAttribute(typeof(AssemblyFileVersionAttribute))).Version.ToString());
                log.Verbose("Telemetry setup done");
            }
            catch (Exception ex)
            {
                this.telemetryClient = null;
                log.Error($"Telemetry setup failed: {ex.Message}. Continuing without telemetry", ex);
            }
        }
        #endregion

        public void LogTransformationStart()
        {
            if (this.telemetryClient == null)
            {
                return;
            }

            try
            {
                // Prepare event data
                Dictionary<string, string> properties = new Dictionary<string, string>();
                Dictionary<string, double> metrics = new Dictionary<string, double>();

                this.telemetryClient.TrackEvent("TransformationService.PageStart", properties, metrics);
            }
            catch
            {
                // Eat all exceptions 
            }
        }

        public void LogTransformationDone(TimeSpan duration)
        {
            if (this.telemetryClient == null)
            {
                return;
            }

            try
            {
                // Prepare event data
                Dictionary<string, string> properties = new Dictionary<string, string>();
                Dictionary<string, double> metrics = new Dictionary<string, double>();

                if (duration != null)
                {
                    properties.Add("Duration", duration.Seconds.ToString());
                }

                this.telemetryClient.TrackEvent("TransformationService.PageDone", properties, metrics);

                // Also add to the metric of transformed pages via the service endpoint
                this.telemetryClient.GetMetric($"TransformationService.PagesTransformed").TrackValue(1);
                this.telemetryClient.GetMetric($"TransformationService.PageDuration").TrackValue(duration.TotalSeconds);
            }
            catch
            {
                // Eat all exceptions 
            }
        }

        public void LogError(Exception ex, string location)
        {
            if (this.telemetryClient == null || ex == null)
            {
                return;
            }

            try
            {
                // Prepare event data
                Dictionary<string, string> properties = new Dictionary<string, string>();
                Dictionary<string, double> metrics = new Dictionary<string, double>();

                if (!string.IsNullOrEmpty(location))
                {
                    properties.Add("Location", location);
                }

                this.telemetryClient.TrackException(ex, properties, metrics);
            }
            catch (Exception ex2)
            {
                // Eat all exceptions 
            }
        }

        /// <summary>
        /// Ensure telemetry data is send to server
        /// </summary>
        public void Flush()
        {
            try
            {
                // before exit, flush the remaining data
                this.telemetryClient.Flush();

                // flush is not blocking so wait a bit
                Task.Delay(50).Wait();
            }
            catch
            {
                // Eat all exceptions
            }
        }
    }
}
