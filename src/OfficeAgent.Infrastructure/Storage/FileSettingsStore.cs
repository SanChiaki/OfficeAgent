using System;
using System.IO;
using System.Security.Cryptography;
using Newtonsoft.Json;
using OfficeAgent.Core.Models;
using OfficeAgent.Infrastructure.Security;

namespace OfficeAgent.Infrastructure.Storage
{
    public sealed class FileSettingsStore
    {
        private readonly string settingsPath;
        private readonly DpapiSecretProtector secretProtector;

        public FileSettingsStore(string settingsPath, DpapiSecretProtector secretProtector)
        {
            this.settingsPath = settingsPath ?? throw new ArgumentNullException(nameof(settingsPath));
            this.secretProtector = secretProtector ?? throw new ArgumentNullException(nameof(secretProtector));
        }

        public AppSettings Load()
        {
            if (!File.Exists(settingsPath))
            {
                return new AppSettings();
            }

            try
            {
                var persisted = JsonConvert.DeserializeObject<PersistedSettings>(File.ReadAllText(settingsPath));
                if (persisted == null)
                {
                    return new AppSettings();
                }

                var settings = new AppSettings
                {
                    ApiKey = string.Empty,
                    BaseUrl = AppSettings.NormalizeBaseUrl(persisted.BaseUrl),
                    BusinessBaseUrl = AppSettings.NormalizeOptionalUrl(persisted.BusinessBaseUrl),
                    Model = string.IsNullOrWhiteSpace(persisted.Model) ? "gpt-5-mini" : persisted.Model,
                    ApiFormat = AppSettings.NormalizeApiFormat(persisted.ApiFormat),
                    UiLanguageOverride = AppSettings.NormalizeUiLanguageOverride(persisted.UiLanguageOverride),
                    SsoUrl = persisted.SsoUrl ?? string.Empty,
                    SsoLoginSuccessPath = persisted.SsoLoginSuccessPath ?? string.Empty,
                };

                try
                {
                    settings.ApiKey = secretProtector.Unprotect(persisted.EncryptedApiKey);
                }
                catch (FormatException)
                {
                    settings.ApiKey = string.Empty;
                }
                catch (CryptographicException)
                {
                    settings.ApiKey = string.Empty;
                }

                return settings;
            }
            catch (JsonException)
            {
                return new AppSettings();
            }
        }

        public void Save(AppSettings settings)
        {
            var persisted = new PersistedSettings
            {
                EncryptedApiKey = secretProtector.Protect(settings?.ApiKey ?? string.Empty),
                BaseUrl = AppSettings.NormalizeBaseUrl(settings?.BaseUrl),
                BusinessBaseUrl = AppSettings.NormalizeOptionalUrl(settings?.BusinessBaseUrl),
                Model = string.IsNullOrWhiteSpace(settings?.Model) ? "gpt-5-mini" : settings.Model,
                ApiFormat = AppSettings.NormalizeApiFormat(settings?.ApiFormat),
                UiLanguageOverride = AppSettings.NormalizeUiLanguageOverride(settings?.UiLanguageOverride),
                SsoUrl = settings?.SsoUrl ?? string.Empty,
                SsoLoginSuccessPath = settings?.SsoLoginSuccessPath ?? string.Empty,
            };

            var directoryPath = Path.GetDirectoryName(settingsPath);
            if (!string.IsNullOrEmpty(directoryPath))
            {
                Directory.CreateDirectory(directoryPath);
            }

            File.WriteAllText(settingsPath, JsonConvert.SerializeObject(persisted, Formatting.Indented));
        }

        private sealed class PersistedSettings
        {
            public string EncryptedApiKey { get; set; } = string.Empty;

            public string BaseUrl { get; set; } = string.Empty;

            public string BusinessBaseUrl { get; set; } = string.Empty;

            public string Model { get; set; } = string.Empty;

            public string ApiFormat { get; set; } = string.Empty;

            public string UiLanguageOverride { get; set; } = string.Empty;

            public string SsoUrl { get; set; } = string.Empty;

            public string SsoLoginSuccessPath { get; set; } = string.Empty;
        }
    }
}
