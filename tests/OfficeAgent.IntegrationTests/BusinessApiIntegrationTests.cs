using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using OfficeAgent.Core.Models;
using OfficeAgent.Infrastructure.Http;
using OfficeAgent.Infrastructure.Storage;
using Xunit;
using FileSettingsStore = OfficeAgent.Infrastructure.Storage.FileSettingsStore;
using DpapiSecretProtector = OfficeAgent.Infrastructure.Security.DpapiSecretProtector;

namespace OfficeAgent.IntegrationTests
{
    public sealed class BusinessApiIntegrationTests : IClassFixture<MockServerFixture>
    {
        private const string ProjectA = "\u9879\u76EEA";

        private readonly MockServerFixture fixture;

        public BusinessApiIntegrationTests(MockServerFixture fixture)
        {
            this.fixture = fixture;
        }

        [Fact]
        public async Task UploadsDataToMockServerAndReceivesSavedCount()
        {
            var cookieJar = await fixture.LoginAs("test_user", "password123");
            var settingsStore = CreateSettingsStore($"{fixture.BusinessUrl}");
            var client = new BusinessApiClient(
                () => settingsStore.Load(),
                new System.Net.Http.HttpClient(new System.Net.Http.HttpClientHandler { UseCookies = true, CookieContainer = cookieJar })
                {
                    Timeout = TimeSpan.FromSeconds(10),
                },
                cookieJar);

            var preview = CreatePreview();
            var result = client.Upload(preview);

            Assert.Equal(2, result.SavedCount);
            Assert.Contains("2", result.Message);
        }

        [Fact]
        public void UploadWithoutAuthReturns401()
        {
            var client = new BusinessApiClient(
                () => new AppSettings { BusinessBaseUrl = fixture.BusinessUrl, ApiKey = string.Empty },
                CreateHttpClientWithoutCookies());

            var ex = Assert.Throws<InvalidOperationException>(() => client.Upload(CreatePreview()));

            Assert.Contains("401", ex.Message);
        }

        [Fact]
        public async Task UploadRetriesAndSucceedsAfterMockServerError()
        {
            var cookieJar = await fixture.LoginAs("retry_user", "password123");
            var settingsStore = CreateSettingsStore($"{fixture.BusinessUrl}");
            var client = new BusinessApiClient(
                () => settingsStore.Load(),
                new System.Net.Http.HttpClient(new System.Net.Http.HttpClientHandler { UseCookies = true, CookieContainer = cookieJar })
                {
                    Timeout = TimeSpan.FromSeconds(10),
                },
                cookieJar);

            var preview = CreatePreview();
            var result = client.Upload(preview);

            Assert.Equal(2, result.SavedCount);
        }

        private static FileSettingsStore CreateSettingsStore(string baseUrl)
        {
            var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + "_settings.json");
            var store = new FileSettingsStore(path, new DpapiSecretProtector());
            store.Save(new AppSettings { BusinessBaseUrl = baseUrl, ApiKey = string.Empty });
            return store;
        }

        private static System.Net.Http.HttpClient CreateHttpClientWithoutCookies()
        {
            var handler = new System.Net.Http.HttpClientHandler { UseCookies = false };
            return new System.Net.Http.HttpClient(handler) { Timeout = TimeSpan.FromSeconds(10) };
        }

        private static UploadPreview CreatePreview()
        {
            return new UploadPreview
            {
                ProjectName = ProjectA,
                SheetName = "Sheet1",
                Address = "A1:C3",
                Headers = new[] { "Name", "Region" },
                Rows = new[]
                {
                    new[] { "Project A", "CN" },
                    new[] { "Project B", "US" },
                },
                Records = new[]
                {
                    new Dictionary<string, string> { ["Name"] = "Project A", ["Region"] = "CN" },
                    new Dictionary<string, string> { ["Name"] = "Project B", ["Region"] = "US" },
                },
            };
        }
    }

    public sealed class MockServerFixture : IDisposable
    {
        public readonly int SsoPort = 3100;
        public readonly int BusinessPort = 3200;
        public readonly string SsoUrl = "http://localhost:3100";
        public readonly string BusinessUrl = "http://localhost:3200";

        private readonly Process process;

        public MockServerFixture()
        {
            var scriptPath = FindMockServerScript(AppDomain.CurrentDomain.BaseDirectory);
            if (scriptPath == null)
            {
                throw new FileNotFoundException("Mock server script not found after walking up from the test assembly directory and current working directory.");
            }

            process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = "node",
                    Arguments = $"\"{scriptPath}\"",
                    WorkingDirectory = Path.GetDirectoryName(scriptPath),
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                },
            };
            process.Start();

            using var client = new System.Net.Http.HttpClient { Timeout = TimeSpan.FromSeconds(1) };
            for (int i = 0; i < 30; i++)
            {
                Thread.Sleep(200);
                try
                {
                    client.GetAsync(SsoUrl + "/login").GetAwaiter().GetResult().EnsureSuccessStatusCode();
                    return;
                }
                catch { }
            }

            throw new InvalidOperationException("Mock server did not start within the timeout.");
        }

        public void Dispose()
        {
            if (process != null && !process.HasExited)
            {
                process.Kill();
                process.WaitForExit(2000);
            }
        }

        /// <summary>
        /// Performs SSO login against the mock server and returns the cookie container.
        /// </summary>
        public async Task<CookieContainer> LoginAs(string username, string password)
        {
            var container = new CookieContainer();
            var handler = new HttpClientHandler
            {
                UseCookies = true,
                CookieContainer = container,
            };
            using var client = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(5) };

            var formData = new FormUrlEncodedContent(new[]
            {
                new KeyValuePair<string, string>("username", username),
                new KeyValuePair<string, string>("password", password),
            });
            using var response = await client.PostAsync(SsoUrl + "/rest/login", formData);
            response.EnsureSuccessStatusCode();
            return container;
        }

        private static string FindMockServerScript(string startDir)
        {
            if (startDir == null) return null;
            var dir = startDir.TrimEnd(Path.DirectorySeparatorChar);
            for (int i = 0; i < 10; i++)
            {
                var candidate = Path.Combine(dir, "tests", "mock-server", "server.js");
                if (File.Exists(candidate)) return candidate;
                var parent = Path.GetDirectoryName(dir);
                if (parent == dir) break;
                dir = parent;
            }
            return null;
        }
    }
}
