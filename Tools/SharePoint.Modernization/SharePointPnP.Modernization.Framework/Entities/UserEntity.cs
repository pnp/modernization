using Newtonsoft.Json;

namespace SharePointPnP.Modernization.Framework.Entities
{
    public class UserEntity
    {
        [JsonProperty(PropertyName = "id")]
        public string Id { get; set; }

        [JsonProperty(PropertyName = "upn")]
        public string Upn { get; set; }

        [JsonProperty(PropertyName = "name")]
        public string Name { get; set; }

        [JsonProperty(PropertyName = "role")]
        public string Role { get; set; }
    }
}
