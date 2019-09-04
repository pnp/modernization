using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.Extensibility;
using Microsoft.SharePoint.Client;
using SharePointPnP.Modernization.Framework.Transform;
using System;
using System.Collections.Generic;
using System.Net;

namespace SharePointPnP.Modernization.Framework.Telemetry
{
    public class PageTelemetry
    {
        private readonly TelemetryClient telemetryClient;
        private Guid tenantId;

        #region Construction
        /// <summary>
        /// Instantiates the telemetry client
        /// </summary>
        public PageTelemetry(string version)
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
                this.telemetryClient.Context.GlobalProperties.Add("Version", version);
            }
            catch (Exception ex)
            {
                this.telemetryClient = null;
            }
        }
        #endregion

        public void LogTransformationDone(TimeSpan duration, string pageType, BaseTransformationInformation baseTransformationInformation)
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

                this.telemetryClient.TrackEvent("TransformationEngine.PageDone", properties, metrics);

                // Also add to the metric of transformed pages via the service endpoint
                this.telemetryClient.GetMetric($"TransformationEngine.PagesTransformed").TrackValue(1);
                this.telemetryClient.GetMetric($"TransformationEngine.PageDuration").TrackValue(duration.TotalSeconds);
                this.telemetryClient.GetMetric($"TransformationEngine.{pageType}").TrackValue(1);

                // Log source environment type 
                this.telemetryClient.GetMetric($"TransformationEngine.Source{baseTransformationInformation.SourceVersion.ToString()}").TrackValue(1);

                // Cross farm or not?
                if (baseTransformationInformation.IsCrossFarmTransformation)
                {
                    this.telemetryClient.GetMetric($"TransformationEngine.CrossFarmTransformation").TrackValue(1);
                }
                else
                {
                    this.telemetryClient.GetMetric($"TransformationEngine.IntraFarmTransformation").TrackValue(1);
                }
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
            }
            catch
            {
                // Eat all exceptions
            }
        }

        #region Helper methods
        internal void LoadAADTenantId(ClientContext context)
        {
            WebRequest request = WebRequest.Create(new Uri(context.Web.Url) + "/_vti_bin/client.svc");
            request.Headers.Add("Authorization: Bearer ");

            try
            {
                using (request.GetResponse())
                {
                }
            }
            catch (WebException e)
            {
                var bearerResponseHeader = e.Response.Headers["WWW-Authenticate"];

                const string bearer = "Bearer realm=\"";
                var bearerIndex = bearerResponseHeader.IndexOf(bearer, StringComparison.Ordinal);

                var realmIndex = bearerIndex + bearer.Length;

                if (bearerResponseHeader.Length >= realmIndex + 36)
                {
                    var targetRealm = bearerResponseHeader.Substring(realmIndex, 36);

                    Guid realmGuid;

                    if (Guid.TryParse(targetRealm, out realmGuid))
                    {
                        this.tenantId = realmGuid;

                        if (!this.telemetryClient.Context.GlobalProperties.ContainsKey("AADTenantId"))
                        {
                            this.telemetryClient.Context.GlobalProperties.Add("AADTenantId", this.tenantId.ToString());
                        }
                    }
                }
            }
        }
        #endregion
    }
}
