using System;
using System.IO;
using OfficeAgent.Infrastructure.Security;
using OfficeAgent.Infrastructure.Storage;
using Xunit;

namespace OfficeAgent.Infrastructure.Tests
{
    public sealed class FileSettingsStoreTests : IDisposable
    {
        private readonly string tempDirectory;

        public FileSettingsStoreTests()
        {
            tempDirectory = Path.Combine(Path.GetTempPath(), "OfficeAgent.Tests", Guid.NewGuid().ToString("N"));
        }

        [Fact]
        public void LoadReturnsDefaultsWhenSettingsFileIsMissing()
        {
            var store = new FileSettingsStore(Path.Combine(tempDirectory, "settings.json"), new DpapiSecretProtector());

            var settings = store.Load();

            Assert.Equal(string.Empty, settings.ApiKey);
            Assert.Equal("https://api.example.com", settings.BaseUrl);
            Assert.Equal(string.Empty, settings.BusinessBaseUrl);
            Assert.Equal("gpt-5-mini", settings.Model);
            Assert.Equal("openai-compatible", settings.ApiFormat);
            Assert.Equal("system", settings.UiLanguageOverride);
        }

        [Fact]
        public void SaveRoundTripsProtectedApiKeyAndUiLanguageOverride()
        {
            var settingsPath = Path.Combine(tempDirectory, "settings.json");
            var store = new FileSettingsStore(settingsPath, new DpapiSecretProtector());

            store.Save(new OfficeAgent.Core.Models.AppSettings
            {
                ApiKey = "secret-token",
                BaseUrl = "https://api.internal.example",
                BusinessBaseUrl = "https://business.internal.example",
                Model = "gpt-5-mini",
                ApiFormat = "anthropic-messages",
                UiLanguageOverride = "zh",
            });

            var persistedJson = File.ReadAllText(settingsPath);
            var loaded = store.Load();

            Assert.DoesNotContain("secret-token", persistedJson);
            Assert.Equal("secret-token", loaded.ApiKey);
            Assert.Equal("https://api.internal.example", loaded.BaseUrl);
            Assert.Equal("https://business.internal.example", loaded.BusinessBaseUrl);
            Assert.Equal("anthropic-messages", loaded.ApiFormat);
            Assert.Equal("zh", loaded.UiLanguageOverride);
            Assert.Contains("\"ApiFormat\": \"anthropic-messages\"", persistedJson);
            Assert.Contains("\"UiLanguageOverride\": \"zh\"", persistedJson);
        }

        [Fact]
        public void SaveNormalizesBaseUrlByTrimmingWhitespaceAndTrailingSlashes()
        {
            var settingsPath = Path.Combine(tempDirectory, "settings.json");
            var store = new FileSettingsStore(settingsPath, new DpapiSecretProtector());

            store.Save(new OfficeAgent.Core.Models.AppSettings
            {
                ApiKey = "secret-token",
                BaseUrl = " https://api.internal.example/// ",
                BusinessBaseUrl = " https://business.internal.example/// ",
                Model = "gpt-5-mini",
            });

            var loaded = store.Load();

            Assert.Equal("https://api.internal.example", loaded.BaseUrl);
            Assert.Equal("https://business.internal.example", loaded.BusinessBaseUrl);
        }

        [Fact]
        public void SaveRoundTripsAnalyticsBaseUrl()
        {
            var settingsPath = Path.Combine(tempDirectory, "settings.json");
            var store = new FileSettingsStore(settingsPath, new DpapiSecretProtector());

            store.Save(new OfficeAgent.Core.Models.AppSettings
            {
                AnalyticsBaseUrl = " https://analytics.internal.example/// ",
            });

            var loaded = store.Load();
            var persistedJson = File.ReadAllText(settingsPath);

            Assert.Equal("https://analytics.internal.example", loaded.AnalyticsBaseUrl);
            Assert.Contains("\"AnalyticsBaseUrl\": \"https://analytics.internal.example\"", persistedJson);
        }

        [Fact]
        public void LoadDefaultsAnalyticsBaseUrlToEmptyString()
        {
            var store = new FileSettingsStore(Path.Combine(tempDirectory, "missing-settings.json"), new DpapiSecretProtector());

            var settings = store.Load();

            Assert.Equal(string.Empty, settings.AnalyticsBaseUrl);
        }

        [Fact]
        public void LoadNormalizesInvalidUiLanguageOverrideToSystem()
        {
            var settingsPath = Path.Combine(tempDirectory, "settings.json");
            Directory.CreateDirectory(tempDirectory);
            File.WriteAllText(
                settingsPath,
                "{\n  \"encryptedApiKey\": \"\",\n  \"baseUrl\": \"https://api.internal.example\",\n  \"model\": \"gpt-5-mini\",\n  \"uiLanguageOverride\": \"de\"\n}");
            var store = new FileSettingsStore(settingsPath, new DpapiSecretProtector());

            var settings = store.Load();

            Assert.Equal("system", settings.UiLanguageOverride);
        }

        [Fact]
        public void LoadNormalizesInvalidApiFormatToOpenAiCompatible()
        {
            var settingsPath = Path.Combine(tempDirectory, "settings.json");
            Directory.CreateDirectory(tempDirectory);
            File.WriteAllText(
                settingsPath,
                "{\n  \"encryptedApiKey\": \"\",\n  \"baseUrl\": \"https://api.internal.example\",\n  \"model\": \"gpt-5-mini\",\n  \"apiFormat\": \"unknown\"\n}");
            var store = new FileSettingsStore(settingsPath, new DpapiSecretProtector());

            var settings = store.Load();

            Assert.Equal("openai-compatible", settings.ApiFormat);
        }

        [Fact]
        public void SaveNormalizesInvalidApiFormatToOpenAiCompatible()
        {
            var settingsPath = Path.Combine(tempDirectory, "settings.json");
            var store = new FileSettingsStore(settingsPath, new DpapiSecretProtector());

            store.Save(new OfficeAgent.Core.Models.AppSettings
            {
                ApiKey = "secret-token",
                ApiFormat = "unknown",
            });

            var loaded = store.Load();

            Assert.Equal("openai-compatible", loaded.ApiFormat);
        }

        [Fact]
        public void SaveNormalizesInvalidUiLanguageOverrideToSystem()
        {
            var settingsPath = Path.Combine(tempDirectory, "settings.json");
            var store = new FileSettingsStore(settingsPath, new DpapiSecretProtector());

            store.Save(new OfficeAgent.Core.Models.AppSettings
            {
                ApiKey = "secret-token",
                UiLanguageOverride = "de",
            });

            var loaded = store.Load();

            Assert.Equal("system", loaded.UiLanguageOverride);
        }

        [Fact]
        public void LoadRecoversWhenProtectedApiKeyCannotBeDecrypted()
        {
            var settingsPath = Path.Combine(tempDirectory, "settings.json");
            Directory.CreateDirectory(tempDirectory);
            File.WriteAllText(
                settingsPath,
                "{\n  \"encryptedApiKey\": \"not-base64\",\n  \"baseUrl\": \"https://api.internal.example\",\n  \"model\": \"gpt-5-mini\"\n}");
            var store = new FileSettingsStore(settingsPath, new DpapiSecretProtector());

            var settings = store.Load();

            Assert.Equal(string.Empty, settings.ApiKey);
            Assert.Equal("https://api.internal.example", settings.BaseUrl);
            Assert.Equal(string.Empty, settings.BusinessBaseUrl);
            Assert.Equal("gpt-5-mini", settings.Model);
        }

        public void Dispose()
        {
            if (Directory.Exists(tempDirectory))
            {
                Directory.Delete(tempDirectory, recursive: true);
            }
        }
    }
}
