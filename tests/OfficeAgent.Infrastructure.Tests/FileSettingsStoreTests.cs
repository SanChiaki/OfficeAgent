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
        }

        [Fact]
        public void SaveRoundTripsProtectedApiKey()
        {
            var settingsPath = Path.Combine(tempDirectory, "settings.json");
            var store = new FileSettingsStore(settingsPath, new DpapiSecretProtector());

            store.Save(new OfficeAgent.Core.Models.AppSettings
            {
                ApiKey = "secret-token",
                BaseUrl = "https://api.internal.example",
                BusinessBaseUrl = "https://business.internal.example",
                Model = "gpt-5-mini",
            });

            var persistedJson = File.ReadAllText(settingsPath);
            var loaded = store.Load();

            Assert.DoesNotContain("secret-token", persistedJson);
            Assert.Equal("secret-token", loaded.ApiKey);
            Assert.Equal("https://api.internal.example", loaded.BaseUrl);
            Assert.Equal("https://business.internal.example", loaded.BusinessBaseUrl);
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
