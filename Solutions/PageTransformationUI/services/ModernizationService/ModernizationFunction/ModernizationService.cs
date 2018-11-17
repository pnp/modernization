using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using SharePointPnP.Modernization.Framework.Transform;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Threading;
using SharePointPnP.ModernizationFunction.Telemetry;

namespace SharePointPnP.ModernizationFunction
{
    /// <summary>
    /// Class holding our Azure modernization functions
    /// </summary>
    public static class ModernizationService
    {
        private static readonly string aadInstance = "https://login.microsoftonline.com/";
        private static FunctionTelemetry telemetry;

        [FunctionName("ModernizePage")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            DateTime transformationStartDateTime = DateTime.Now;
            // instantiate the telemetry model
            telemetry = new FunctionTelemetry(log);
            if (telemetry != null)
            {
                telemetry.LogTransformationStart();
            }

            try
            {
                log.Info("ModernizePage Azure function is being called");

                string transformedPageUrl = null;

                // Parse and process the query parameters
                string siteUrl = req.GetQueryNameValuePairs().FirstOrDefault(q => string.Compare(q.Key, "SiteUrl", true) == 0).Value;
                string pageUrl = req.GetQueryNameValuePairs().FirstOrDefault(q => string.Compare(q.Key, "PageUrl", true) == 0).Value;

                if (string.IsNullOrEmpty(siteUrl) || string.IsNullOrEmpty(pageUrl))
                {
                    throw new Exception($"Invalid page and/or site url. PageUrl:{pageUrl} SiteUrl:{siteUrl}");
                }

                Uri siteUri = new Uri(siteUrl);

                // is this request authenticated?
                if (Thread.CurrentPrincipal != null &&
                    Thread.CurrentPrincipal.Identity != null &&
                    Thread.CurrentPrincipal.Identity.IsAuthenticated)
                {
                    var issuer = (Thread.CurrentPrincipal as ClaimsPrincipal)?.FindFirst("iss");
                    if (issuer != null && !String.IsNullOrEmpty(issuer.Value))
                    {
                        var issuerValue = issuer.Value.Substring(0, issuer.Value.Length - 1);
                        var tenantId = issuerValue.Substring(issuerValue.LastIndexOf("/") + 1);
                        var upn = (Thread.CurrentPrincipal as ClaimsPrincipal)?.FindFirst(ClaimTypes.Upn)?.Value;

                        // Check for multitenant usage of the service. Used for demo purposes or for customers that want to host a 
                        // single service for multiple tenants
                        if (!IsAllowedTenant(upn))
                        {                            
                            throw new Exception($"Tenant {upn.Substring(upn.IndexOf('@') + 1).ToLower()} is not whitelisted for this transformation service endpoint.");
                        }

                        string accessToken = "";
                        // We need to manually retrieve an access token valid for SharePoint from the 
                        // received access token used to access this service
                        using (var client = new HttpClient())
                        {
                            // Prepare the AAD OAuth request URI
                            string authority = $"{aadInstance}common";
                            var tokenUri = new Uri($"{authority}/oauth2/token");

                            // Grab the access token that was in the request (the one coming from SPFX)
                            log.Verbose($"Parameter: {req.Headers.Authorization.Parameter}");

                            // Prepare the OAuth 2.0 request for an Access Token with the on behalf of flow
                            // https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow
                            var content = new FormUrlEncodedContent(new[]
                            {
                                new KeyValuePair<string, string>("grant_type", "urn:ietf:params:oauth:grant-type:jwt-bearer"),
                                new KeyValuePair<string, string>("client_id", GetAppSetting("CLIENT_ID")),
                                new KeyValuePair<string, string>("client_secret", GetAppSetting("CLIENT_SECRET")),
                                new KeyValuePair<string, string>("assertion", req.Headers.Authorization.Parameter),
                                new KeyValuePair<string, string>("resource", $"https://{siteUri.Authority}"),
                                new KeyValuePair<string, string>("requested_token_use", "on_behalf_of"),
                            });

                            // Make the HTTP request
                            var result = await client.PostAsync(tokenUri, content);
                            string jsonToken = await result.Content.ReadAsStringAsync();

                            // Get back the OAuth 2.0 response
                            var token = JsonConvert.DeserializeObject<OAuthTokenResponse>(jsonToken);

                            if (token != null && !String.IsNullOrEmpty(token.AccessToken))
                            {
                                accessToken = token.AccessToken;
                            }
                            else
                            {
                                log.Error($"Obtaining an access token for SharePoint failed. Request response = {jsonToken}");
                                return req.CreateResponse(HttpStatusCode.Unauthorized, jsonToken);
                            }
                        }

                        log.Info($"IssuerValue: {issuerValue} tenantId: {tenantId} upn: {upn}");
                        log.Verbose($"accessToken: {accessToken}");

                        if (!string.IsNullOrEmpty(accessToken))
                        {
                            try
                            {
                                log.Info($"Transforming page {pageUrl}");
                                OfficeDevPnP.Core.AuthenticationManager authManager = new OfficeDevPnP.Core.AuthenticationManager();
                                using (ClientContext context = authManager.GetAzureADAccessTokenAuthenticatedContext(siteUrl, accessToken))
                                {
                                    string fileName = System.IO.Path.GetFileName(pageUrl);
                                    var pages = context.Web.GetPageToTransform(fileName);

                                    if (pages.Count == 1)
                                    {
                                        log.Verbose($"Page {fileName} was found and will be transformed");

                                        // Load page transformation settings 
                                        bool replaceHomePageWithDefaultHomePage = GetAppSettingBool("ReplaceHomePageWithDefaultHomePage", false);
                                        string targetPagePrefix = GetAppSetting("TargetPagePrefix", "Migrated_");

                                        log.Verbose($"Options used to drive page transformation: TargetPagePrefix={targetPagePrefix} ReplaceHomePageWithDefaultHomePage={replaceHomePageWithDefaultHomePage}");

                                        var pageTransformator = new PageTransformator(context, "d:\\home\\site\\wwwroot\\webpartmapping.xml");
                                        PageTransformationInformation pti = new PageTransformationInformation(pages[0])
                                        {
                                            // If target page exists, then overwrite it
                                            Overwrite = true,

                                            // Modernization center setup
                                            ModernizationCenterInformation = new ModernizationCenterInformation()
                                            {
                                                AddPageAcceptBanner = true,
                                            },

                                            // Give the migrated page a specific prefix, default is Migrated_
                                            TargetPagePrefix = targetPagePrefix,

                                            // If the page is a home page then replace with stock home page
                                            ReplaceHomePageWithDefaultHomePage = replaceHomePageWithDefaultHomePage,
                                        };

                                        pageTransformator.Transform(pti);

                                        if (pti.TargetPageTakesSourcePageName)
                                        {
                                            transformedPageUrl = pageUrl;
                                        }
                                        else
                                        {
                                            transformedPageUrl = pageUrl.Replace(fileName, $"{pti.TargetPagePrefix}{fileName}");
                                        }

                                        log.Info($"Page {fileName} was transformed into {transformedPageUrl}");

                                        TimeSpan duration = DateTime.Now.Subtract(transformationStartDateTime);
                                        if (telemetry != null)
                                        {
                                            telemetry.LogTransformationDone(duration);
                                        }

                                        return req.CreateResponse(HttpStatusCode.OK, $"{transformedPageUrl}");
                                    }
                                    else
                                    {
                                        throw new Exception($"This function only supports the transformation of a single page. We found {pages.Count} pages for page {fileName}");
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                if (telemetry != null)
                                {
                                    telemetry.LogError(ex, "TransformationLoop");
                                }
                                log.Error(ex.ToDetailedString());
                                return req.CreateResponse(HttpStatusCode.InternalServerError, ex.Message);
                            }
                        }
                        else
                        {
                            throw new Exception("Access token could not be retrieved");
                        }
                    }
                    else
                    {
                        throw new Exception("Issuer value could not be retrieved");
                    }
                }
                else
                {
                    throw new Exception("Request was not having a user identity. Are you using this function in anonymous mode?");
                }
            }
            catch (Exception ex)
            {
                if (telemetry != null)
                {
                    telemetry.LogError(ex, "MainLoop");
                }
                log.Error(ex.ToDetailedString());
                return req.CreateResponse(HttpStatusCode.BadRequest, ex.Message);
            }
            finally
            {
                if (telemetry != null)
                {
                    telemetry.Flush();
                }
            }
        }

        private static bool IsAllowedTenant(string upn)
        {
            string tenants = GetAppSetting("AllowedTenants");
            
            // no specific config set...
            if (string.IsNullOrEmpty(tenants) || string.IsNullOrEmpty(upn))
            {
                return true;
            }

            var tenantName = upn.Substring(upn.IndexOf('@') + 1).ToLower();
            string[] tenantNames = tenants.ToLower().Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            return tenantNames.Contains(tenantName);
        }


        private static string GetAppSetting(string name)
        {
            return Environment.GetEnvironmentVariable(name, EnvironmentVariableTarget.Process);
        }

        private static string GetAppSetting(string name, string defaultValue)
        {
            var result = Environment.GetEnvironmentVariable(name, EnvironmentVariableTarget.Process);
            if (result == null)
            {
                return defaultValue;
            }
            else
            {
                return result;
            }
        }

        private static bool GetAppSettingBool(string name, bool defaultValue)
        {
            if (bool.TryParse(Environment.GetEnvironmentVariable(name, EnvironmentVariableTarget.Process), out bool result))
            {
                return result;
            }
            return defaultValue;
        }

    }
}
