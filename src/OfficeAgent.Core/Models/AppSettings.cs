namespace OfficeAgent.Core.Models
{
    public sealed class AppSettings
    {
        public const string DefaultBaseUrl = "https://api.example.com";

        public string ApiKey { get; set; } = string.Empty;

        public string BaseUrl { get; set; } = DefaultBaseUrl;

        public string Model { get; set; } = "gpt-5-mini";

        public string SsoUrl { get; set; } = string.Empty;

        public static string NormalizeBaseUrl(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return DefaultBaseUrl;
            }

            var normalized = value.Trim().TrimEnd('/');
            return string.IsNullOrWhiteSpace(normalized) ? DefaultBaseUrl : normalized;
        }
    }
}
