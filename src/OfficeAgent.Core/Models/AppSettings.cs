using System;

namespace OfficeAgent.Core.Models
{
    public sealed class AppSettings
    {
        public const string DefaultBaseUrl = "https://api.example.com";
        public const string DefaultApiFormat = "openai-compatible";
        public const string DefaultUiLanguageOverride = "system";

        public string ApiKey { get; set; } = string.Empty;

        public string BaseUrl { get; set; } = DefaultBaseUrl;

        public string BusinessBaseUrl { get; set; } = string.Empty;

        public string Model { get; set; } = "gpt-5-mini";

        public string ApiFormat { get; set; } = DefaultApiFormat;

        public string UiLanguageOverride { get; set; } = DefaultUiLanguageOverride;

        public string SsoUrl { get; set; } = string.Empty;

        public string SsoLoginSuccessPath { get; set; } = string.Empty;

        public static string NormalizeBaseUrl(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return DefaultBaseUrl;
            }

            var normalized = value.Trim().TrimEnd('/');
            return string.IsNullOrWhiteSpace(normalized) ? DefaultBaseUrl : normalized;
        }

        public static string NormalizeOptionalUrl(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            return value.Trim().TrimEnd('/');
        }

        public static string NormalizeUiLanguageOverride(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return DefaultUiLanguageOverride;
            }

            var normalized = value.Trim();
            if (string.Equals(normalized, "system", StringComparison.OrdinalIgnoreCase))
            {
                return DefaultUiLanguageOverride;
            }

            if (string.Equals(normalized, "zh", StringComparison.OrdinalIgnoreCase))
            {
                return "zh";
            }

            if (string.Equals(normalized, "en", StringComparison.OrdinalIgnoreCase))
            {
                return "en";
            }

            return DefaultUiLanguageOverride;
        }

        public static string NormalizeApiFormat(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return DefaultApiFormat;
            }

            var normalized = value.Trim();
            if (string.Equals(normalized, DefaultApiFormat, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(normalized, "openai", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(normalized, "chat-completions", StringComparison.OrdinalIgnoreCase))
            {
                return DefaultApiFormat;
            }

            if (string.Equals(normalized, "anthropic-messages", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(normalized, "anthropic", StringComparison.OrdinalIgnoreCase))
            {
                return "anthropic-messages";
            }

            return DefaultApiFormat;
        }
    }
}
