using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;

namespace OfficeAgent.Core.Models
{
    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public sealed class FetchResult
    {
        public bool Success { get; set; }

        public int StatusCode { get; set; }

        public string Body { get; set; } = string.Empty;

        public string ErrorMessage { get; set; }
    }
}
