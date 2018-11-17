using Newtonsoft.Json;
using System;

namespace SharePointPnP.ModernizationFunction
{

    public class OAuthTokenResponse
    {
        [JsonProperty("token_type")]
        public String TokenType { get; set; }

        [JsonProperty("scope")]
        public String Scope { get; set; }

        [JsonProperty("expires_in")]
        public Int64 ExpiresIn { get; set; }

        [JsonProperty("ext_expires_in")]
        public Int64 ExtExpiresIn { get; set; }

        [JsonProperty("expires_on")]
        public Int64 ExpiresOn { get; set; }

        [JsonProperty("not_before")]
        public Int64 NotBefore { get; set; }

        [JsonProperty("resource")]
        public String Resource { get; set; }

        [JsonProperty("access_token")]
        public String AccessToken { get; set; }

        [JsonProperty("refresh_token")]
        public String RefreshToken { get; set; }

        [JsonProperty("id_token")]
        public String IdToken { get; set; }
    }
}
