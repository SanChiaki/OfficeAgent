# Ribbon Version Reminder Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a Release-only, non-blocking new-version reminder that shows a red dot on the Ribbon `About` button, reads an independent JSON update manifest, caches update state, and lets users ignore the current latest version.

**Architecture:** Keep update-check business logic in focused ExcelAddIn `Updates` classes, keep Ribbon code limited to binding state and displaying UI, and keep the task pane settings untouched. `ThisAddIn` composes the update service, loads cached state synchronously, then schedules a delayed background check without awaiting it. The update service owns Release/config gating, cache windows, version comparison, ignore state, and failure isolation.

**Tech Stack:** C# .NET Framework 4.8, VSTO Ribbon, WinForms, `HttpClient`, Newtonsoft.Json, xUnit, existing `OfficeAgentLog`.

---

## File Structure

- Create `src/OfficeAgent.ExcelAddIn/Updates/UpdateVersionComparer.cs`: parse and compare current vs latest versions.
- Create `src/OfficeAgent.ExcelAddIn/Updates/UpdateManifest.cs`: JSON DTO returned by the update source.
- Create `src/OfficeAgent.ExcelAddIn/Updates/UpdateState.cs`: persisted `%LocalAppData%\OfficeAgent\update-state.json` state.
- Create `src/OfficeAgent.ExcelAddIn/Updates/UpdateNotificationState.cs`: immutable state consumed by Ribbon.
- Create `src/OfficeAgent.ExcelAddIn/Updates/UpdateCheckOptions.cs`: internal options for enablement, URL, cache duration, timeout, and startup delay.
- Create `src/OfficeAgent.ExcelAddIn/Updates/UpdateCheckConfiguration.cs`: Release-only default configuration and registry override lookup.
- Create `src/OfficeAgent.ExcelAddIn/Updates/IUpdateManifestClient.cs`: manifest HTTP client contract.
- Create `src/OfficeAgent.ExcelAddIn/Updates/UpdateManifestClient.cs`: GET update JSON, tolerate `application/octet-stream`, validate fields.
- Create `src/OfficeAgent.ExcelAddIn/Updates/IUpdateStateStore.cs`: update state persistence contract.
- Create `src/OfficeAgent.ExcelAddIn/Updates/FileUpdateStateStore.cs`: safe JSON file persistence.
- Create `src/OfficeAgent.ExcelAddIn/Updates/UpdateNotificationService.cs`: cache, ignore, background scheduling, and non-blocking state change notifications.
- Create `src/OfficeAgent.ExcelAddIn/Dialogs/AboutDialog.cs`: WinForms About dialog with update links and Ignore action.
- Create `src/OfficeAgent.ExcelAddIn/RibbonAboutIconFactory.cs`: generate normal and red-dot About button images at runtime.
- Modify `src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj`: include the new source files.
- Modify `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs`: compose update service and start the delayed background check.
- Modify `src/OfficeAgent.ExcelAddIn/AgentRibbon.cs`: bind update state, refresh About icon, and route Ignore action.
- Modify `src/OfficeAgent.ExcelAddIn/Localization/HostLocalizedStrings.cs`: add About update strings.
- Modify `tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs`: source-level integration checks.
- Create `tests/OfficeAgent.ExcelAddIn.Tests/UpdateVersionComparerTests.cs`.
- Create `tests/OfficeAgent.ExcelAddIn.Tests/UpdateManifestClientTests.cs`.
- Create `tests/OfficeAgent.ExcelAddIn.Tests/FileUpdateStateStoreTests.cs`.
- Create `tests/OfficeAgent.ExcelAddIn.Tests/UpdateNotificationServiceTests.cs`.
- Modify `docs/modules/ribbon-sync-current-behavior.md`.
- Modify `docs/vsto-manual-test-checklist.md`.
- Modify `docs/ribbon-button-custom-icons-guide.md`.

---

### Task 1: Version DTOs And Comparison

**Files:**
- Create: `src/OfficeAgent.ExcelAddIn/Updates/UpdateVersionComparer.cs`
- Create: `src/OfficeAgent.ExcelAddIn/Updates/UpdateManifest.cs`
- Create: `src/OfficeAgent.ExcelAddIn/Updates/UpdateState.cs`
- Create: `src/OfficeAgent.ExcelAddIn/Updates/UpdateNotificationState.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj`
- Test: `tests/OfficeAgent.ExcelAddIn.Tests/UpdateVersionComparerTests.cs`

- [ ] **Step 1: Write the failing version comparison tests**

Create `tests/OfficeAgent.ExcelAddIn.Tests/UpdateVersionComparerTests.cs`:

```csharp
using OfficeAgent.ExcelAddIn.Updates;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class UpdateVersionComparerTests
    {
        [Theory]
        [InlineData("1.0.176", "1.0.175", true)]
        [InlineData("v1.0.176", "1.0.175", true)]
        [InlineData("1.1.0", "1.0.999", true)]
        [InlineData("2.0.0", "1.9.999", true)]
        [InlineData("1.0.175", "1.0.175", false)]
        [InlineData("1.0.174", "1.0.175", false)]
        [InlineData("not-a-version", "1.0.175", false)]
        [InlineData("1.0.176", "not-a-version", false)]
        [InlineData("", "1.0.175", false)]
        [InlineData(null, "1.0.175", false)]
        public void IsNewerThanCurrentComparesSupportedVersions(string latestVersion, string currentVersion, bool expected)
        {
            Assert.Equal(expected, UpdateVersionComparer.IsNewerThanCurrent(latestVersion, currentVersion));
        }
    }
}
```

- [ ] **Step 2: Run the focused tests and verify they fail**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~UpdateVersionComparerTests"
```

Expected: FAIL because `OfficeAgent.ExcelAddIn.Updates.UpdateVersionComparer` does not exist.

- [ ] **Step 3: Add update DTOs and version comparer**

Create `src/OfficeAgent.ExcelAddIn/Updates/UpdateVersionComparer.cs`:

```csharp
using System;

namespace OfficeAgent.ExcelAddIn.Updates
{
    internal static class UpdateVersionComparer
    {
        public static bool IsNewerThanCurrent(string latestVersion, string currentVersion)
        {
            return TryParse(latestVersion, out var latest) &&
                   TryParse(currentVersion, out var current) &&
                   latest.CompareTo(current) > 0;
        }

        public static bool TryParse(string value, out Version version)
        {
            version = null;
            var normalized = (value ?? string.Empty).Trim();
            if (normalized.StartsWith("v", StringComparison.OrdinalIgnoreCase))
            {
                normalized = normalized.Substring(1);
            }

            return Version.TryParse(normalized, out version);
        }
    }
}
```

Create `src/OfficeAgent.ExcelAddIn/Updates/UpdateManifest.cs`:

```csharp
using System;

namespace OfficeAgent.ExcelAddIn.Updates
{
    internal sealed class UpdateManifest
    {
        public string LatestVersion { get; set; } = string.Empty;

        public string DownloadUrl { get; set; } = string.Empty;

        public string ReleaseNotesUrl { get; set; } = string.Empty;

        public DateTime? PublishedAtUtc { get; set; }

        public string Title { get; set; } = string.Empty;

        public string Summary { get; set; } = string.Empty;
    }
}
```

Create `src/OfficeAgent.ExcelAddIn/Updates/UpdateState.cs`:

```csharp
using System;

namespace OfficeAgent.ExcelAddIn.Updates
{
    internal sealed class UpdateState
    {
        public DateTime? LastCheckedAtUtc { get; set; }

        public string LatestVersion { get; set; } = string.Empty;

        public string DownloadUrl { get; set; } = string.Empty;

        public string ReleaseNotesUrl { get; set; } = string.Empty;

        public DateTime? PublishedAtUtc { get; set; }

        public string Title { get; set; } = string.Empty;

        public string Summary { get; set; } = string.Empty;

        public string IgnoredVersion { get; set; } = string.Empty;

        public UpdateManifest ToManifest()
        {
            return new UpdateManifest
            {
                LatestVersion = LatestVersion ?? string.Empty,
                DownloadUrl = DownloadUrl ?? string.Empty,
                ReleaseNotesUrl = ReleaseNotesUrl ?? string.Empty,
                PublishedAtUtc = PublishedAtUtc,
                Title = Title ?? string.Empty,
                Summary = Summary ?? string.Empty,
            };
        }

        public void ApplyManifest(UpdateManifest manifest, DateTime checkedAtUtc)
        {
            LastCheckedAtUtc = checkedAtUtc;
            LatestVersion = manifest?.LatestVersion ?? string.Empty;
            DownloadUrl = manifest?.DownloadUrl ?? string.Empty;
            ReleaseNotesUrl = manifest?.ReleaseNotesUrl ?? string.Empty;
            PublishedAtUtc = manifest?.PublishedAtUtc;
            Title = manifest?.Title ?? string.Empty;
            Summary = manifest?.Summary ?? string.Empty;
        }
    }
}
```

Create `src/OfficeAgent.ExcelAddIn/Updates/UpdateNotificationState.cs`:

```csharp
using System;

namespace OfficeAgent.ExcelAddIn.Updates
{
    internal sealed class UpdateNotificationState
    {
        public static readonly UpdateNotificationState Empty = new UpdateNotificationState(
            hasNewVersion: false,
            latestVersion: string.Empty,
            downloadUrl: string.Empty,
            releaseNotesUrl: string.Empty,
            publishedAtUtc: null,
            title: string.Empty,
            summary: string.Empty);

        public UpdateNotificationState(
            bool hasNewVersion,
            string latestVersion,
            string downloadUrl,
            string releaseNotesUrl,
            DateTime? publishedAtUtc,
            string title,
            string summary)
        {
            HasNewVersion = hasNewVersion;
            LatestVersion = latestVersion ?? string.Empty;
            DownloadUrl = downloadUrl ?? string.Empty;
            ReleaseNotesUrl = releaseNotesUrl ?? string.Empty;
            PublishedAtUtc = publishedAtUtc;
            Title = title ?? string.Empty;
            Summary = summary ?? string.Empty;
        }

        public bool HasNewVersion { get; }

        public string LatestVersion { get; }

        public string DownloadUrl { get; }

        public string ReleaseNotesUrl { get; }

        public DateTime? PublishedAtUtc { get; }

        public string Title { get; }

        public string Summary { get; }
    }
}
```

Modify `src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj` in the main `<ItemGroup>` of `<Compile Include=...>` entries by adding:

```xml
    <Compile Include="Updates\UpdateVersionComparer.cs" />
    <Compile Include="Updates\UpdateManifest.cs" />
    <Compile Include="Updates\UpdateState.cs" />
    <Compile Include="Updates\UpdateNotificationState.cs" />
```

- [ ] **Step 4: Run version tests and verify they pass**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~UpdateVersionComparerTests"
```

Expected: PASS.

- [ ] **Step 5: Commit Task 1**

```powershell
git add src/OfficeAgent.ExcelAddIn/Updates/UpdateVersionComparer.cs src/OfficeAgent.ExcelAddIn/Updates/UpdateManifest.cs src/OfficeAgent.ExcelAddIn/Updates/UpdateState.cs src/OfficeAgent.ExcelAddIn/Updates/UpdateNotificationState.cs src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj tests/OfficeAgent.ExcelAddIn.Tests/UpdateVersionComparerTests.cs
git commit -m "feat: add update version model"
```

---

### Task 2: Manifest Client And State Store

**Files:**
- Create: `src/OfficeAgent.ExcelAddIn/Updates/IUpdateManifestClient.cs`
- Create: `src/OfficeAgent.ExcelAddIn/Updates/UpdateManifestClient.cs`
- Create: `src/OfficeAgent.ExcelAddIn/Updates/IUpdateStateStore.cs`
- Create: `src/OfficeAgent.ExcelAddIn/Updates/FileUpdateStateStore.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj`
- Test: `tests/OfficeAgent.ExcelAddIn.Tests/UpdateManifestClientTests.cs`
- Test: `tests/OfficeAgent.ExcelAddIn.Tests/FileUpdateStateStoreTests.cs`

- [ ] **Step 1: Write failing manifest client tests**

Create `tests/OfficeAgent.ExcelAddIn.Tests/UpdateManifestClientTests.cs`:

```csharp
using System;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeAgent.ExcelAddIn.Updates;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class UpdateManifestClientTests
    {
        [Fact]
        public async Task GetManifestAsyncParsesJsonEvenWhenContentTypeIsOctetStream()
        {
            const string json = "{\"latestVersion\":\"1.0.176\",\"downloadUrl\":\"https://updates.example/download.exe\",\"releaseNotesUrl\":\"https://updates.example/notes\",\"publishedAtUtc\":\"2026-05-19T08:00:00Z\",\"title\":\"Release\",\"summary\":\"Summary\"}";
            var handler = new StubHttpHandler(_ =>
            {
                var response = new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ByteArrayContent(Encoding.UTF8.GetBytes(json)),
                };
                response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                return response;
            });
            var client = new UpdateManifestClient(new HttpClient(handler));

            var manifest = await client.GetManifestAsync("https://updates.example/manifest", CancellationToken.None);

            Assert.Equal("1.0.176", manifest.LatestVersion);
            Assert.Equal("https://updates.example/download.exe", manifest.DownloadUrl);
            Assert.Equal("https://updates.example/notes", manifest.ReleaseNotesUrl);
            Assert.Equal(new DateTime(2026, 5, 19, 8, 0, 0, DateTimeKind.Utc), manifest.PublishedAtUtc);
            Assert.Equal("Release", manifest.Title);
            Assert.Equal("Summary", manifest.Summary);
            Assert.Equal(1, handler.Calls);
        }

        [Theory]
        [InlineData("{\"downloadUrl\":\"https://updates.example/download.exe\"}")]
        [InlineData("{\"latestVersion\":\"\"}")]
        [InlineData("not-json")]
        public async Task GetManifestAsyncRejectsInvalidManifests(string body)
        {
            var handler = new StubHttpHandler(_ => new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(body, Encoding.UTF8, "application/octet-stream"),
            });
            var client = new UpdateManifestClient(new HttpClient(handler));

            await Assert.ThrowsAsync<InvalidOperationException>(
                () => client.GetManifestAsync("https://updates.example/manifest", CancellationToken.None));
        }

        [Fact]
        public async Task GetManifestAsyncRejectsNonSuccessResponses()
        {
            var client = new UpdateManifestClient(new HttpClient(new StubHttpHandler(_ =>
                new HttpResponseMessage(HttpStatusCode.InternalServerError))));

            await Assert.ThrowsAsync<HttpRequestException>(
                () => client.GetManifestAsync("https://updates.example/manifest", CancellationToken.None));
        }

        private sealed class StubHttpHandler : HttpMessageHandler
        {
            private readonly Func<HttpRequestMessage, HttpResponseMessage> createResponse;

            public StubHttpHandler(Func<HttpRequestMessage, HttpResponseMessage> createResponse)
            {
                this.createResponse = createResponse;
            }

            public int Calls { get; private set; }

            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
            {
                Calls++;
                return Task.FromResult(createResponse(request));
            }
        }
    }
}
```

- [ ] **Step 2: Write failing state store tests**

Create `tests/OfficeAgent.ExcelAddIn.Tests/FileUpdateStateStoreTests.cs`:

```csharp
using System;
using System.IO;
using OfficeAgent.ExcelAddIn.Updates;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class FileUpdateStateStoreTests
    {
        [Fact]
        public void SaveAndLoadRoundTripsUpdateState()
        {
            var path = CreateTempPath();
            try
            {
                var store = new FileUpdateStateStore(path);
                var state = new UpdateState
                {
                    LastCheckedAtUtc = new DateTime(2026, 5, 19, 8, 0, 0, DateTimeKind.Utc),
                    LatestVersion = "1.0.176",
                    DownloadUrl = "https://updates.example/download.exe",
                    ReleaseNotesUrl = "https://updates.example/notes",
                    PublishedAtUtc = new DateTime(2026, 5, 19, 7, 0, 0, DateTimeKind.Utc),
                    Title = "Release",
                    Summary = "Summary",
                    IgnoredVersion = "1.0.175",
                };

                store.Save(state);
                var loaded = store.Load();

                Assert.Equal(state.LastCheckedAtUtc, loaded.LastCheckedAtUtc);
                Assert.Equal("1.0.176", loaded.LatestVersion);
                Assert.Equal("https://updates.example/download.exe", loaded.DownloadUrl);
                Assert.Equal("https://updates.example/notes", loaded.ReleaseNotesUrl);
                Assert.Equal(state.PublishedAtUtc, loaded.PublishedAtUtc);
                Assert.Equal("Release", loaded.Title);
                Assert.Equal("Summary", loaded.Summary);
                Assert.Equal("1.0.175", loaded.IgnoredVersion);
            }
            finally
            {
                DeleteIfExists(path);
            }
        }

        [Fact]
        public void LoadReturnsEmptyStateWhenFileIsMissingOrCorrupt()
        {
            var missingPath = CreateTempPath();
            var corruptPath = CreateTempPath();
            try
            {
                File.WriteAllText(corruptPath, "{ invalid json");

                Assert.Equal(string.Empty, new FileUpdateStateStore(missingPath).Load().LatestVersion);
                Assert.Equal(string.Empty, new FileUpdateStateStore(corruptPath).Load().LatestVersion);
            }
            finally
            {
                DeleteIfExists(missingPath);
                DeleteIfExists(corruptPath);
            }
        }

        private static string CreateTempPath()
        {
            return Path.Combine(Path.GetTempPath(), "officeagent-update-state-" + Guid.NewGuid().ToString("N") + ".json");
        }

        private static void DeleteIfExists(string path)
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }
}
```

- [ ] **Step 3: Run the focused tests and verify they fail**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~UpdateManifestClientTests|FullyQualifiedName~FileUpdateStateStoreTests"
```

Expected: FAIL because the manifest client and state store types do not exist.

- [ ] **Step 4: Add manifest client and state store implementation**

Create `src/OfficeAgent.ExcelAddIn/Updates/IUpdateManifestClient.cs`:

```csharp
using System.Threading;
using System.Threading.Tasks;

namespace OfficeAgent.ExcelAddIn.Updates
{
    internal interface IUpdateManifestClient
    {
        Task<UpdateManifest> GetManifestAsync(string manifestUrl, CancellationToken cancellationToken);
    }
}
```

Create `src/OfficeAgent.ExcelAddIn/Updates/UpdateManifestClient.cs`:

```csharp
using System;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace OfficeAgent.ExcelAddIn.Updates
{
    internal sealed class UpdateManifestClient : IUpdateManifestClient
    {
        private readonly HttpClient httpClient;

        public UpdateManifestClient(HttpClient httpClient = null)
        {
            this.httpClient = httpClient ?? new HttpClient();
        }

        public async Task<UpdateManifest> GetManifestAsync(string manifestUrl, CancellationToken cancellationToken)
        {
            if (!Uri.TryCreate(manifestUrl, UriKind.Absolute, out var uri) ||
                (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps))
            {
                throw new InvalidOperationException("Update manifest URL must be an absolute HTTP or HTTPS URL.");
            }

            using (var request = new HttpRequestMessage(HttpMethod.Get, uri))
            using (var response = await httpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false))
            {
                if (!response.IsSuccessStatusCode)
                {
                    throw new HttpRequestException($"Update manifest request failed with HTTP {(int)response.StatusCode}.");
                }

                var body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                UpdateManifest manifest;
                try
                {
                    manifest = JsonConvert.DeserializeObject<UpdateManifest>(body);
                }
                catch (JsonException ex)
                {
                    throw new InvalidOperationException("Update manifest response was not valid JSON.", ex);
                }

                if (manifest == null || string.IsNullOrWhiteSpace(manifest.LatestVersion))
                {
                    throw new InvalidOperationException("Update manifest is missing latestVersion.");
                }

                manifest.DownloadUrl = NormalizeOptionalHttpUrl(manifest.DownloadUrl);
                manifest.ReleaseNotesUrl = NormalizeOptionalHttpUrl(manifest.ReleaseNotesUrl);
                manifest.Title = manifest.Title ?? string.Empty;
                manifest.Summary = manifest.Summary ?? string.Empty;
                return manifest;
            }
        }

        private static string NormalizeOptionalHttpUrl(string value)
        {
            var normalized = (value ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return string.Empty;
            }

            return Uri.TryCreate(normalized, UriKind.Absolute, out var uri) &&
                   (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps)
                ? normalized
                : string.Empty;
        }
    }
}
```

Create `src/OfficeAgent.ExcelAddIn/Updates/IUpdateStateStore.cs`:

```csharp
namespace OfficeAgent.ExcelAddIn.Updates
{
    internal interface IUpdateStateStore
    {
        UpdateState Load();

        void Save(UpdateState state);
    }
}
```

Create `src/OfficeAgent.ExcelAddIn/Updates/FileUpdateStateStore.cs`:

```csharp
using System;
using System.IO;
using Newtonsoft.Json;
using OfficeAgent.Core.Diagnostics;

namespace OfficeAgent.ExcelAddIn.Updates
{
    internal sealed class FileUpdateStateStore : IUpdateStateStore
    {
        private readonly string path;

        public FileUpdateStateStore(string path)
        {
            this.path = path ?? throw new ArgumentNullException(nameof(path));
        }

        public UpdateState Load()
        {
            try
            {
                if (!File.Exists(path))
                {
                    return new UpdateState();
                }

                var json = File.ReadAllText(path);
                return JsonConvert.DeserializeObject<UpdateState>(json) ?? new UpdateState();
            }
            catch (Exception ex)
            {
                OfficeAgentLog.Warn("updates", "state.load_failed", $"Failed to load update state from {path}. {ex.Message}");
                return new UpdateState();
            }
        }

        public void Save(UpdateState state)
        {
            try
            {
                var directory = Path.GetDirectoryName(path);
                if (!string.IsNullOrWhiteSpace(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                File.WriteAllText(path, JsonConvert.SerializeObject(state ?? new UpdateState(), Formatting.Indented));
            }
            catch (Exception ex)
            {
                OfficeAgentLog.Error("updates", "state.save_failed", $"Failed to save update state to {path}.", ex);
            }
        }
    }
}
```

Modify `src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj` by adding:

```xml
    <Compile Include="Updates\IUpdateManifestClient.cs" />
    <Compile Include="Updates\UpdateManifestClient.cs" />
    <Compile Include="Updates\IUpdateStateStore.cs" />
    <Compile Include="Updates\FileUpdateStateStore.cs" />
```

- [ ] **Step 5: Run manifest and state store tests**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~UpdateManifestClientTests|FullyQualifiedName~FileUpdateStateStoreTests"
```

Expected: PASS.

- [ ] **Step 6: Commit Task 2**

```powershell
git add src/OfficeAgent.ExcelAddIn/Updates/IUpdateManifestClient.cs src/OfficeAgent.ExcelAddIn/Updates/UpdateManifestClient.cs src/OfficeAgent.ExcelAddIn/Updates/IUpdateStateStore.cs src/OfficeAgent.ExcelAddIn/Updates/FileUpdateStateStore.cs src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj tests/OfficeAgent.ExcelAddIn.Tests/UpdateManifestClientTests.cs tests/OfficeAgent.ExcelAddIn.Tests/FileUpdateStateStoreTests.cs
git commit -m "feat: add update manifest persistence"
```

---

### Task 3: Update Notification Service

**Files:**
- Create: `src/OfficeAgent.ExcelAddIn/Updates/UpdateCheckOptions.cs`
- Create: `src/OfficeAgent.ExcelAddIn/Updates/UpdateCheckConfiguration.cs`
- Create: `src/OfficeAgent.ExcelAddIn/Updates/UpdateNotificationService.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj`
- Test: `tests/OfficeAgent.ExcelAddIn.Tests/UpdateNotificationServiceTests.cs`

- [ ] **Step 1: Write failing service behavior tests**

Create `tests/OfficeAgent.ExcelAddIn.Tests/UpdateNotificationServiceTests.cs`:

```csharp
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using OfficeAgent.ExcelAddIn.Updates;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class UpdateNotificationServiceTests
    {
        [Fact]
        public async Task CheckForUpdatesAsyncSkipsHttpWhenDisabled()
        {
            var client = new FakeManifestClient(new UpdateManifest { LatestVersion = "1.0.176" });
            var store = new MemoryUpdateStateStore(new UpdateState());
            var service = CreateService(UpdateCheckOptions.Disabled(), client, store);

            await service.CheckForUpdatesAsync(CancellationToken.None);

            Assert.Equal(0, client.Calls);
            Assert.False(service.CurrentState.HasNewVersion);
        }

        [Fact]
        public async Task CheckForUpdatesAsyncUsesCacheWindowWithoutHttp()
        {
            var now = new DateTime(2026, 5, 19, 8, 0, 0, DateTimeKind.Utc);
            var client = new FakeManifestClient(new UpdateManifest { LatestVersion = "1.0.177" });
            var store = new MemoryUpdateStateStore(new UpdateState
            {
                LastCheckedAtUtc = now.AddHours(-1),
                LatestVersion = "1.0.176",
            });
            var service = CreateService(UpdateCheckOptions.Enabled("https://updates.example/manifest"), client, store, now);

            await service.CheckForUpdatesAsync(CancellationToken.None);

            Assert.Equal(0, client.Calls);
            Assert.True(service.CurrentState.HasNewVersion);
            Assert.Equal("1.0.176", service.CurrentState.LatestVersion);
        }

        [Fact]
        public async Task CheckForUpdatesAsyncShowsNewVersionWhenLatestIsHigherAndNotIgnored()
        {
            var now = new DateTime(2026, 5, 19, 8, 0, 0, DateTimeKind.Utc);
            var manifest = new UpdateManifest
            {
                LatestVersion = "1.0.176",
                DownloadUrl = "https://updates.example/download.exe",
                ReleaseNotesUrl = "https://updates.example/notes",
                PublishedAtUtc = now,
                Title = "Release",
                Summary = "Summary",
            };
            var client = new FakeManifestClient(manifest);
            var store = new MemoryUpdateStateStore(new UpdateState());
            var service = CreateService(UpdateCheckOptions.Enabled("https://updates.example/manifest"), client, store, now);

            await service.CheckForUpdatesAsync(CancellationToken.None);

            Assert.Equal(1, client.Calls);
            Assert.True(service.CurrentState.HasNewVersion);
            Assert.Equal("1.0.176", service.CurrentState.LatestVersion);
            Assert.Equal("https://updates.example/download.exe", service.CurrentState.DownloadUrl);
            Assert.Equal("1.0.176", store.State.LatestVersion);
            Assert.Equal(now, store.State.LastCheckedAtUtc);
        }

        [Fact]
        public async Task CheckForUpdatesAsyncHidesIgnoredVersionButShowsHigherFutureVersion()
        {
            var now = new DateTime(2026, 5, 19, 8, 0, 0, DateTimeKind.Utc);
            var store = new MemoryUpdateStateStore(new UpdateState { IgnoredVersion = "1.0.176" });
            var service = CreateService(
                UpdateCheckOptions.Enabled("https://updates.example/manifest"),
                new FakeManifestClient(new UpdateManifest { LatestVersion = "1.0.176" }),
                store,
                now);

            await service.CheckForUpdatesAsync(CancellationToken.None);
            Assert.False(service.CurrentState.HasNewVersion);

            var nextService = CreateService(
                UpdateCheckOptions.Enabled("https://updates.example/manifest"),
                new FakeManifestClient(new UpdateManifest { LatestVersion = "1.0.177" }),
                store,
                now.AddDays(2));

            await nextService.CheckForUpdatesAsync(CancellationToken.None);
            Assert.True(nextService.CurrentState.HasNewVersion);
            Assert.Equal("1.0.177", nextService.CurrentState.LatestVersion);
        }

        [Fact]
        public async Task IgnoreCurrentVersionPersistsIgnoredVersionAndRaisesStateChanged()
        {
            var store = new MemoryUpdateStateStore(new UpdateState());
            var service = CreateService(
                UpdateCheckOptions.Enabled("https://updates.example/manifest"),
                new FakeManifestClient(new UpdateManifest { LatestVersion = "1.0.176" }),
                store);
            var events = new List<UpdateNotificationState>();
            service.StateChanged += (_, __) => events.Add(service.CurrentState);
            await service.CheckForUpdatesAsync(CancellationToken.None);

            service.IgnoreCurrentVersion();

            Assert.Equal("1.0.176", store.State.IgnoredVersion);
            Assert.False(service.CurrentState.HasNewVersion);
            Assert.True(events.Count >= 2);
        }

        [Fact]
        public void StartBackgroundCheckReturnsBeforeTheDelayedRequestRuns()
        {
            var client = new FakeManifestClient(new UpdateManifest { LatestVersion = "1.0.176" });
            var service = CreateService(
                UpdateCheckOptions.Enabled(
                    "https://updates.example/manifest",
                    cacheDuration: TimeSpan.FromHours(24),
                    requestTimeout: TimeSpan.FromSeconds(5),
                    startupDelay: TimeSpan.FromMilliseconds(200)),
                client,
                new MemoryUpdateStateStore(new UpdateState()));

            service.StartBackgroundCheck(uiContext: null);

            Assert.Equal(0, client.Calls);
        }

        private static UpdateNotificationService CreateService(
            UpdateCheckOptions options,
            IUpdateManifestClient client,
            IUpdateStateStore store,
            DateTime? now = null)
        {
            return new UpdateNotificationService(
                options,
                client,
                store,
                currentVersion: "1.0.175",
                getUtcNow: () => now ?? new DateTime(2026, 5, 19, 8, 0, 0, DateTimeKind.Utc));
        }

        private sealed class FakeManifestClient : IUpdateManifestClient
        {
            private readonly UpdateManifest manifest;

            public FakeManifestClient(UpdateManifest manifest)
            {
                this.manifest = manifest;
            }

            public int Calls { get; private set; }

            public Task<UpdateManifest> GetManifestAsync(string manifestUrl, CancellationToken cancellationToken)
            {
                Calls++;
                return Task.FromResult(manifest);
            }
        }

        private sealed class MemoryUpdateStateStore : IUpdateStateStore
        {
            public MemoryUpdateStateStore(UpdateState state)
            {
                State = state;
            }

            public UpdateState State { get; private set; }

            public UpdateState Load()
            {
                return State;
            }

            public void Save(UpdateState state)
            {
                State = state;
            }
        }
    }
}
```

- [ ] **Step 2: Run service tests and verify they fail**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~UpdateNotificationServiceTests"
```

Expected: FAIL because `UpdateNotificationService` and `UpdateCheckOptions` do not exist.

- [ ] **Step 3: Add options, configuration, and service**

Create `src/OfficeAgent.ExcelAddIn/Updates/UpdateCheckOptions.cs`:

```csharp
using System;

namespace OfficeAgent.ExcelAddIn.Updates
{
    internal sealed class UpdateCheckOptions
    {
        private UpdateCheckOptions(
            bool isEnabled,
            string manifestUrl,
            TimeSpan cacheDuration,
            TimeSpan requestTimeout,
            TimeSpan startupDelay)
        {
            IsEnabled = isEnabled;
            ManifestUrl = manifestUrl ?? string.Empty;
            CacheDuration = cacheDuration;
            RequestTimeout = requestTimeout;
            StartupDelay = startupDelay;
        }

        public bool IsEnabled { get; }

        public string ManifestUrl { get; }

        public TimeSpan CacheDuration { get; }

        public TimeSpan RequestTimeout { get; }

        public TimeSpan StartupDelay { get; }

        public static UpdateCheckOptions Enabled(
            string manifestUrl,
            TimeSpan? cacheDuration = null,
            TimeSpan? requestTimeout = null,
            TimeSpan? startupDelay = null)
        {
            return new UpdateCheckOptions(
                isEnabled: true,
                manifestUrl: manifestUrl,
                cacheDuration: cacheDuration ?? TimeSpan.FromHours(24),
                requestTimeout: requestTimeout ?? TimeSpan.FromSeconds(5),
                startupDelay: startupDelay ?? TimeSpan.FromSeconds(5));
        }

        public static UpdateCheckOptions Disabled()
        {
            return new UpdateCheckOptions(
                isEnabled: false,
                manifestUrl: string.Empty,
                cacheDuration: TimeSpan.FromHours(24),
                requestTimeout: TimeSpan.FromSeconds(5),
                startupDelay: TimeSpan.FromSeconds(5));
        }
    }
}
```

Create `src/OfficeAgent.ExcelAddIn/Updates/UpdateCheckConfiguration.cs`:

```csharp
using System;
using Microsoft.Win32;
using OfficeAgent.Core.Diagnostics;

namespace OfficeAgent.ExcelAddIn.Updates
{
    internal static class UpdateCheckConfiguration
    {
        private const string RegistrySubKey = @"Software\OfficeAgent";
        private const string RegistryValueName = "UpdateManifestUrl";
        private const string DefaultManifestUrl = "";

        public static UpdateCheckOptions CreateDefault()
        {
#if DEBUG
            OfficeAgentLog.Info("updates", "configuration.disabled_debug", "Update checks are disabled for Debug builds.");
            return UpdateCheckOptions.Disabled();
#else
            var manifestUrl = ReadRegistryManifestUrl();
            if (string.IsNullOrWhiteSpace(manifestUrl))
            {
                manifestUrl = DefaultManifestUrl;
            }

            if (string.IsNullOrWhiteSpace(manifestUrl))
            {
                OfficeAgentLog.Info("updates", "configuration.disabled_missing_url", "Update checks are disabled because no manifest URL is configured.");
                return UpdateCheckOptions.Disabled();
            }

            return UpdateCheckOptions.Enabled(manifestUrl);
#endif
        }

        private static string ReadRegistryManifestUrl()
        {
            return ReadRegistryManifestUrl(Registry.CurrentUser) ??
                   ReadRegistryManifestUrl(Registry.LocalMachine) ??
                   string.Empty;
        }

        private static string ReadRegistryManifestUrl(RegistryKey root)
        {
            try
            {
                using (var key = root.OpenSubKey(RegistrySubKey))
                {
                    return key?.GetValue(RegistryValueName) as string;
                }
            }
            catch (Exception ex)
            {
                OfficeAgentLog.Warn("updates", "configuration.registry_read_failed", $"Failed to read update manifest registry value. {ex.Message}");
                return string.Empty;
            }
        }
    }
}
```

Create `src/OfficeAgent.ExcelAddIn/Updates/UpdateNotificationService.cs`:

```csharp
using System;
using System.Threading;
using System.Threading.Tasks;
using OfficeAgent.Core.Diagnostics;

namespace OfficeAgent.ExcelAddIn.Updates
{
    internal sealed class UpdateNotificationService
    {
        private readonly UpdateCheckOptions options;
        private readonly IUpdateManifestClient manifestClient;
        private readonly IUpdateStateStore stateStore;
        private readonly string currentVersion;
        private readonly Func<DateTime> getUtcNow;
        private readonly object syncRoot = new object();
        private UpdateState cachedState;

        public UpdateNotificationService(
            UpdateCheckOptions options,
            IUpdateManifestClient manifestClient,
            IUpdateStateStore stateStore,
            string currentVersion,
            Func<DateTime> getUtcNow = null)
        {
            this.options = options ?? UpdateCheckOptions.Disabled();
            this.manifestClient = manifestClient ?? throw new ArgumentNullException(nameof(manifestClient));
            this.stateStore = stateStore ?? throw new ArgumentNullException(nameof(stateStore));
            this.currentVersion = currentVersion ?? string.Empty;
            this.getUtcNow = getUtcNow ?? (() => DateTime.UtcNow);
            cachedState = new UpdateState();
            CurrentState = UpdateNotificationState.Empty;
        }

        public event EventHandler StateChanged;

        public UpdateNotificationState CurrentState { get; private set; }

        public void LoadCachedState()
        {
            var state = stateStore.Load() ?? new UpdateState();
            lock (syncRoot)
            {
                cachedState = state;
                CurrentState = BuildNotificationState(state);
            }

            RaiseStateChanged(null);
        }

        public void StartBackgroundCheck(SynchronizationContext uiContext)
        {
            Task.Run(async () =>
            {
                try
                {
                    if (options.StartupDelay > TimeSpan.Zero)
                    {
                        await Task.Delay(options.StartupDelay).ConfigureAwait(false);
                    }

                    await CheckForUpdatesAsync(CancellationToken.None).ConfigureAwait(false);
                    RaiseStateChanged(uiContext);
                }
                catch (Exception ex)
                {
                    OfficeAgentLog.Error("updates", "background_check.failed", "Background update check failed.", ex);
                }
            });
        }

        public async Task CheckForUpdatesAsync(CancellationToken cancellationToken)
        {
            if (!options.IsEnabled)
            {
                OfficeAgentLog.Info("updates", "check.skipped_disabled", "Update check skipped because it is disabled.");
                ApplyState(stateStore.Load() ?? new UpdateState());
                return;
            }

            if (string.IsNullOrWhiteSpace(options.ManifestUrl))
            {
                OfficeAgentLog.Info("updates", "check.skipped_missing_url", "Update check skipped because no manifest URL is configured.");
                ApplyState(stateStore.Load() ?? new UpdateState());
                return;
            }

            var state = stateStore.Load() ?? new UpdateState();
            var now = getUtcNow();
            if (state.LastCheckedAtUtc.HasValue &&
                now - state.LastCheckedAtUtc.Value.ToUniversalTime() < options.CacheDuration)
            {
                OfficeAgentLog.Info("updates", "check.skipped_cache_fresh", "Update check skipped because cached state is fresh.");
                ApplyState(state);
                return;
            }

            try
            {
                OfficeAgentLog.Info("updates", "check.started", "Checking for OfficeAgent updates.");
                using (var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken))
                {
                    timeout.CancelAfter(options.RequestTimeout);
                    var manifest = await manifestClient.GetManifestAsync(options.ManifestUrl, timeout.Token).ConfigureAwait(false);
                    if (!UpdateVersionComparer.TryParse(manifest.LatestVersion, out _))
                    {
                        throw new InvalidOperationException("Update manifest latestVersion is not a supported version.");
                    }

                    state.ApplyManifest(manifest, now);
                    stateStore.Save(state);
                    ApplyState(state);
                    OfficeAgentLog.Info("updates", "check.completed", $"Update check completed. LatestVersion={manifest.LatestVersion}");
                }
            }
            catch (Exception ex)
            {
                OfficeAgentLog.Warn("updates", "check.failed", $"Update check failed. {ex.Message}");
                ApplyState(state);
            }
        }

        public void IgnoreCurrentVersion()
        {
            UpdateState state;
            lock (syncRoot)
            {
                state = cachedState ?? stateStore.Load() ?? new UpdateState();
                if (string.IsNullOrWhiteSpace(state.LatestVersion))
                {
                    return;
                }

                state.IgnoredVersion = state.LatestVersion;
                cachedState = state;
                CurrentState = BuildNotificationState(state);
            }

            stateStore.Save(state);
            OfficeAgentLog.Info("updates", "version.ignored", $"Ignored update version {state.IgnoredVersion}.");
            RaiseStateChanged(null);
        }

        private void ApplyState(UpdateState state)
        {
            lock (syncRoot)
            {
                cachedState = state ?? new UpdateState();
                CurrentState = BuildNotificationState(cachedState);
            }
        }

        private UpdateNotificationState BuildNotificationState(UpdateState state)
        {
            if (state == null ||
                string.IsNullOrWhiteSpace(state.LatestVersion) ||
                !UpdateVersionComparer.IsNewerThanCurrent(state.LatestVersion, currentVersion) ||
                string.Equals(state.LatestVersion, state.IgnoredVersion, StringComparison.OrdinalIgnoreCase))
            {
                return UpdateNotificationState.Empty;
            }

            return new UpdateNotificationState(
                hasNewVersion: true,
                latestVersion: state.LatestVersion,
                downloadUrl: state.DownloadUrl,
                releaseNotesUrl: state.ReleaseNotesUrl,
                publishedAtUtc: state.PublishedAtUtc,
                title: state.Title,
                summary: state.Summary);
        }

        private void RaiseStateChanged(SynchronizationContext uiContext)
        {
            var handler = StateChanged;
            if (handler == null)
            {
                return;
            }

            if (uiContext == null)
            {
                handler(this, EventArgs.Empty);
                return;
            }

            uiContext.Post(_ => handler(this, EventArgs.Empty), null);
        }
    }
}
```

Modify `src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj` by adding:

```xml
    <Compile Include="Updates\UpdateCheckOptions.cs" />
    <Compile Include="Updates\UpdateCheckConfiguration.cs" />
    <Compile Include="Updates\UpdateNotificationService.cs" />
```

- [ ] **Step 4: Run service tests**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~UpdateNotificationServiceTests"
```

Expected: PASS.

- [ ] **Step 5: Commit Task 3**

```powershell
git add src/OfficeAgent.ExcelAddIn/Updates/UpdateCheckOptions.cs src/OfficeAgent.ExcelAddIn/Updates/UpdateCheckConfiguration.cs src/OfficeAgent.ExcelAddIn/Updates/UpdateNotificationService.cs src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj tests/OfficeAgent.ExcelAddIn.Tests/UpdateNotificationServiceTests.cs
git commit -m "feat: add update notification service"
```

---

### Task 4: Compose The Non-Blocking Update Service

**Files:**
- Modify: `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs`
- Test: `tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs`

- [ ] **Step 1: Add failing source-level composition tests**

Append these tests to `tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs` before `ResolveRepositoryPath`:

```csharp
        [Fact]
        public void ThisAddInComposesUpdateNotificationServiceWithoutAwaitingBackgroundCheck()
        {
            var addInText = File.ReadAllText(ResolveRepositoryPath("src", "OfficeAgent.ExcelAddIn", "ThisAddIn.cs"));

            Assert.Contains("internal UpdateNotificationService UpdateNotificationService { get; private set; }", addInText, StringComparison.Ordinal);
            Assert.Contains("UpdateNotificationService = new UpdateNotificationService(", addInText, StringComparison.Ordinal);
            Assert.Contains("UpdateNotificationService.LoadCachedState();", addInText, StringComparison.Ordinal);
            Assert.Contains("UpdateNotificationService.StartBackgroundCheck(startupSynchronizationContext);", addInText, StringComparison.Ordinal);
            Assert.DoesNotContain("await UpdateNotificationService", addInText, StringComparison.Ordinal);
        }

        [Fact]
        public void ExcelAddInProjectIncludesUpdateNotificationSources()
        {
            var projectText = File.ReadAllText(ResolveRepositoryPath("src", "OfficeAgent.ExcelAddIn", "OfficeAgent.ExcelAddIn.csproj"));

            Assert.Contains("<Compile Include=\"Updates\\UpdateCheckConfiguration.cs\" />", projectText, StringComparison.Ordinal);
            Assert.Contains("<Compile Include=\"Updates\\UpdateNotificationService.cs\" />", projectText, StringComparison.Ordinal);
            Assert.Contains("<Compile Include=\"Updates\\UpdateManifestClient.cs\" />", projectText, StringComparison.Ordinal);
            Assert.Contains("<Compile Include=\"Updates\\FileUpdateStateStore.cs\" />", projectText, StringComparison.Ordinal);
        }
```

- [ ] **Step 2: Run the focused source tests and verify they fail**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~ThisAddInComposesUpdateNotificationServiceWithoutAwaitingBackgroundCheck|FullyQualifiedName~ExcelAddInProjectIncludesUpdateNotificationSources"
```

Expected: FAIL because `ThisAddIn` does not compose `UpdateNotificationService`.

- [ ] **Step 3: Compose update service in ThisAddIn**

Modify `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs` usings:

```csharp
using System.Threading;
using OfficeAgent.ExcelAddIn.Updates;
```

Add this property near the other internal services:

```csharp
        internal UpdateNotificationService UpdateNotificationService { get; private set; }
```

Inside `ThisAddIn_Startup`, after `var initialSettings = SettingsStore.Load();`, add:

```csharp
            var startupSynchronizationContext = SynchronizationContext.Current;
```

After `AccountSessionService.ConfigureSsoDomain(initialSettings.SsoUrl);`, add:

```csharp
            UpdateNotificationService = new UpdateNotificationService(
                UpdateCheckConfiguration.CreateDefault(),
                new UpdateManifestClient(),
                new FileUpdateStateStore(Path.Combine(appDataDirectory, "update-state.json")),
                VersionInfo.AppVersion);
            UpdateNotificationService.LoadCachedState();
```

After `Globals.Ribbons.AgentRibbon?.BindToControllersAndRefresh();`, add:

```csharp
            UpdateNotificationService.StartBackgroundCheck(startupSynchronizationContext);
```

- [ ] **Step 4: Run composition tests**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~ThisAddInComposesUpdateNotificationServiceWithoutAwaitingBackgroundCheck|FullyQualifiedName~ExcelAddInProjectIncludesUpdateNotificationSources"
```

Expected: PASS.

- [ ] **Step 5: Commit Task 4**

```powershell
git add src/OfficeAgent.ExcelAddIn/ThisAddIn.cs tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs
git commit -m "feat: compose update notification service"
```

---

### Task 5: Ribbon Red Dot And About Dialog

**Files:**
- Create: `src/OfficeAgent.ExcelAddIn/RibbonAboutIconFactory.cs`
- Create: `src/OfficeAgent.ExcelAddIn/Dialogs/AboutDialog.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj`
- Modify: `src/OfficeAgent.ExcelAddIn/Localization/HostLocalizedStrings.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/AgentRibbon.cs`
- Test: `tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs`

- [ ] **Step 1: Add failing Ribbon integration tests**

Append these tests to `tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs` before `ResolveRepositoryPath`:

```csharp
        [Fact]
        public void AgentRibbonBindsUpdateNotificationServiceAndRefreshesAboutIcon()
        {
            var ribbonText = File.ReadAllText(ResolveRepositoryPath("src", "OfficeAgent.ExcelAddIn", "AgentRibbon.cs"));

            Assert.Contains("TryBindToUpdateNotificationService();", ribbonText, StringComparison.Ordinal);
            Assert.Contains("UpdateNotificationService_StateChanged", ribbonText, StringComparison.Ordinal);
            Assert.Contains("ApplyAboutButtonImage();", ribbonText, StringComparison.Ordinal);
            Assert.Contains("aboutButton.OfficeImageId = string.Empty;", ribbonText, StringComparison.Ordinal);
            Assert.Contains("RibbonAboutIconFactory.CreateAboutIcon(hasUpdate:", ribbonText, StringComparison.Ordinal);
        }

        [Fact]
        public void AboutButtonOpensUpdateAwareDialogAndCanIgnoreVersion()
        {
            var ribbonText = File.ReadAllText(ResolveRepositoryPath("src", "OfficeAgent.ExcelAddIn", "AgentRibbon.cs"));
            var dialogText = File.ReadAllText(ResolveRepositoryPath("src", "OfficeAgent.ExcelAddIn", "Dialogs", "AboutDialog.cs"));

            Assert.Contains("AboutDialog.Show(", ribbonText, StringComparison.Ordinal);
            Assert.Contains("IgnoreCurrentVersion();", ribbonText, StringComparison.Ordinal);
            Assert.Contains("AboutDialogAction.IgnoreVersion", dialogText, StringComparison.Ordinal);
            Assert.Contains("DownloadUrl", dialogText, StringComparison.Ordinal);
            Assert.Contains("ReleaseNotesUrl", dialogText, StringComparison.Ordinal);
        }

        [Fact]
        public void HostLocalizedStringsIncludeUpdateReminderText()
        {
            var stringsText = File.ReadAllText(ResolveRepositoryPath("src", "OfficeAgent.ExcelAddIn", "Localization", "HostLocalizedStrings.cs"));

            Assert.Contains("AboutCurrentVersionLabel", stringsText, StringComparison.Ordinal);
            Assert.Contains("AboutLatestVersionLabel", stringsText, StringComparison.Ordinal);
            Assert.Contains("AboutIgnoreVersionButtonText", stringsText, StringComparison.Ordinal);
            Assert.Contains("AboutDownloadButtonText", stringsText, StringComparison.Ordinal);
            Assert.Contains("AboutReleaseNotesButtonText", stringsText, StringComparison.Ordinal);
        }
```

- [ ] **Step 2: Run the focused Ribbon tests and verify they fail**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~AgentRibbonBindsUpdateNotificationServiceAndRefreshesAboutIcon|FullyQualifiedName~AboutButtonOpensUpdateAwareDialogAndCanIgnoreVersion|FullyQualifiedName~HostLocalizedStringsIncludeUpdateReminderText"
```

Expected: FAIL because Ribbon update binding and About dialog do not exist.

- [ ] **Step 3: Add Ribbon About icon factory**

Create `src/OfficeAgent.ExcelAddIn/RibbonAboutIconFactory.cs`:

```csharp
using System.Drawing;
using System.Drawing.Drawing2D;

namespace OfficeAgent.ExcelAddIn
{
    internal static class RibbonAboutIconFactory
    {
        public static Image CreateAboutIcon(bool hasUpdate)
        {
            var bitmap = new Bitmap(32, 32);
            using (var graphics = Graphics.FromImage(bitmap))
            using (var iconBrush = new SolidBrush(Color.FromArgb(45, 95, 170)))
            using (var textBrush = new SolidBrush(Color.White))
            using (var dotBrush = new SolidBrush(Color.FromArgb(217, 48, 37)))
            using (var dotBorderBrush = new SolidBrush(Color.White))
            using (var font = new Font("Segoe UI", 18, FontStyle.Bold, GraphicsUnit.Pixel))
            {
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                graphics.Clear(Color.Transparent);
                graphics.FillEllipse(iconBrush, 4, 4, 24, 24);
                var textSize = graphics.MeasureString("i", font);
                graphics.DrawString(
                    "i",
                    font,
                    textBrush,
                    16 - (textSize.Width / 2),
                    15 - (textSize.Height / 2));

                if (hasUpdate)
                {
                    graphics.FillEllipse(dotBorderBrush, 20, 2, 10, 10);
                    graphics.FillEllipse(dotBrush, 22, 4, 6, 6);
                }
            }

            return bitmap;
        }
    }
}
```

- [ ] **Step 4: Add update-aware About dialog**

Create `src/OfficeAgent.ExcelAddIn/Dialogs/AboutDialog.cs`:

```csharp
using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using OfficeAgent.ExcelAddIn.Localization;

namespace OfficeAgent.ExcelAddIn.Dialogs
{
    internal enum AboutDialogAction
    {
        Close,
        IgnoreVersion,
    }

    internal sealed class AboutDialogModel
    {
        public string AppVersion { get; set; } = string.Empty;

        public string AssemblyVersion { get; set; } = string.Empty;

        public string BuildConfiguration { get; set; } = string.Empty;

        public string BuildTime { get; set; } = string.Empty;

        public bool HasNewVersion { get; set; }

        public string LatestVersion { get; set; } = string.Empty;

        public string DownloadUrl { get; set; } = string.Empty;

        public string ReleaseNotesUrl { get; set; } = string.Empty;

        public DateTime? PublishedAtUtc { get; set; }

        public string UpdateTitle { get; set; } = string.Empty;

        public string UpdateSummary { get; set; } = string.Empty;
    }

    internal sealed class AboutDialog : Form
    {
        private const int DialogWidth = 460;
        private const int HorizontalPadding = 18;
        private const int ButtonHeight = 28;

        private readonly AboutDialogModel model;
        private readonly HostLocalizedStrings strings;
        private AboutDialogAction action = AboutDialogAction.Close;

        private AboutDialog(AboutDialogModel model, HostLocalizedStrings strings)
        {
            this.model = model ?? new AboutDialogModel();
            this.strings = strings ?? HostLocalizedStrings.ForLocale("en");
            BuildLayout();
        }

        public static AboutDialogAction Show(AboutDialogModel model, HostLocalizedStrings strings)
        {
            var owner = ExcelDialogOwner.FromCurrentApplication();
            using (var dialog = new AboutDialog(model, strings))
            {
                if (owner == null)
                {
                    dialog.ShowDialog();
                }
                else
                {
                    dialog.ShowDialog(owner);
                }

                return dialog.action;
            }
        }

        private void BuildLayout()
        {
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Dpi;
            Text = strings.RibbonAboutDialogTitle;
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;

            var top = 16;
            AddLine("OfficeAgent Excel Add-in", FontStyle.Bold, ref top);
            AddLine($"{strings.AboutCurrentVersionLabel}: {model.AppVersion}", FontStyle.Regular, ref top);
            AddLine($"{strings.AboutAssemblyVersionLabel}: {model.AssemblyVersion}", FontStyle.Regular, ref top);
            AddLine($"{strings.AboutBuildConfigurationLabel}: {model.BuildConfiguration}", FontStyle.Regular, ref top);
            AddLine($"{strings.AboutBuildTimeLabel}: {model.BuildTime}", FontStyle.Regular, ref top);

            top += 8;
            if (model.HasNewVersion)
            {
                AddLine(strings.AboutNewVersionAvailableTitle, FontStyle.Bold, ref top);
                AddLine($"{strings.AboutLatestVersionLabel}: {model.LatestVersion}", FontStyle.Regular, ref top);
                if (model.PublishedAtUtc.HasValue)
                {
                    AddLine($"{strings.AboutPublishedAtLabel}: {model.PublishedAtUtc.Value.ToLocalTime():yyyy-MM-dd HH:mm:ss}", FontStyle.Regular, ref top);
                }

                if (!string.IsNullOrWhiteSpace(model.UpdateTitle))
                {
                    AddLine(model.UpdateTitle, FontStyle.Regular, ref top);
                }

                if (!string.IsNullOrWhiteSpace(model.UpdateSummary))
                {
                    AddWrappedLine(model.UpdateSummary, ref top);
                }
            }
            else
            {
                AddLine(strings.AboutNoUpdateAvailableText, FontStyle.Regular, ref top);
            }

            top += 10;
            AddButtons(top);
            ClientSize = new Size(DialogWidth, top + ButtonHeight + 18);
        }

        private void AddLine(string text, FontStyle style, ref int top)
        {
            var label = new Label
            {
                AutoSize = false,
                Text = text ?? string.Empty,
                Font = new Font(Font, style),
                Bounds = new Rectangle(HorizontalPadding, top, DialogWidth - (HorizontalPadding * 2), 22),
            };
            Controls.Add(label);
            top += 22;
        }

        private void AddWrappedLine(string text, ref int top)
        {
            var label = new Label
            {
                AutoSize = false,
                Text = text ?? string.Empty,
                Bounds = new Rectangle(HorizontalPadding, top, DialogWidth - (HorizontalPadding * 2), 48),
            };
            Controls.Add(label);
            top += 50;
        }

        private void AddButtons(int top)
        {
            var right = DialogWidth - HorizontalPadding;
            var closeButton = CreateButton(strings.CloseButtonText, right - 76, top, 76);
            closeButton.DialogResult = DialogResult.Cancel;
            closeButton.Click += (_, __) =>
            {
                action = AboutDialogAction.Close;
                Close();
            };
            Controls.Add(closeButton);
            right -= 84;

            if (model.HasNewVersion)
            {
                var ignoreButton = CreateButton(strings.AboutIgnoreVersionButtonText, right - 118, top, 118);
                ignoreButton.Click += (_, __) =>
                {
                    action = AboutDialogAction.IgnoreVersion;
                    Close();
                };
                Controls.Add(ignoreButton);
                right -= 126;
            }

            if (!string.IsNullOrWhiteSpace(model.ReleaseNotesUrl))
            {
                var releaseNotesButton = CreateButton(strings.AboutReleaseNotesButtonText, right - 110, top, 110);
                releaseNotesButton.Click += (_, __) => OpenUrl(model.ReleaseNotesUrl);
                Controls.Add(releaseNotesButton);
                right -= 118;
            }

            if (!string.IsNullOrWhiteSpace(model.DownloadUrl))
            {
                var downloadButton = CreateButton(strings.AboutDownloadButtonText, right - 88, top, 88);
                downloadButton.Click += (_, __) => OpenUrl(model.DownloadUrl);
                Controls.Add(downloadButton);
            }

            CancelButton = closeButton;
        }

        private Button CreateButton(string text, int left, int top, int width)
        {
            return new Button
            {
                Text = text ?? string.Empty,
                Bounds = new Rectangle(Math.Max(HorizontalPadding, left), top, width, ButtonHeight),
            };
        }

        private static void OpenUrl(string url)
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = url,
                UseShellExecute = true,
            });
        }
    }
}
```

- [ ] **Step 5: Add localized About update strings**

Modify `src/OfficeAgent.ExcelAddIn/Localization/HostLocalizedStrings.cs` near `RibbonAboutDialogTitle` and `AboutMessage` by adding:

```csharp
        public string AboutCurrentVersionLabel => Locale == "zh" ? "当前版本" : "Current version";

        public string AboutAssemblyVersionLabel => Locale == "zh" ? "程序集版本" : "Assembly version";

        public string AboutBuildConfigurationLabel => Locale == "zh" ? "构建配置" : "Build configuration";

        public string AboutBuildTimeLabel => Locale == "zh" ? "构建时间" : "Build time";

        public string AboutNewVersionAvailableTitle => Locale == "zh" ? "发现新版本" : "New version available";

        public string AboutLatestVersionLabel => Locale == "zh" ? "最新版本" : "Latest version";

        public string AboutPublishedAtLabel => Locale == "zh" ? "发布时间" : "Published";

        public string AboutNoUpdateAvailableText => Locale == "zh" ? "当前没有可用的新版本。" : "No new version is available.";

        public string AboutDownloadButtonText => Locale == "zh" ? "下载" : "Download";

        public string AboutReleaseNotesButtonText => Locale == "zh" ? "发布说明" : "Release notes";

        public string AboutIgnoreVersionButtonText => Locale == "zh" ? "忽略此版本" : "Ignore this version";
```

- [ ] **Step 6: Include new Ribbon and dialog sources**

Modify `src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj` by adding:

```xml
    <Compile Include="RibbonAboutIconFactory.cs" />
    <Compile Include="Dialogs\AboutDialog.cs" />
```

- [ ] **Step 7: Bind update service and About dialog in AgentRibbon**

Modify `src/OfficeAgent.ExcelAddIn/AgentRibbon.cs` usings:

```csharp
using System.Drawing;
using OfficeAgent.ExcelAddIn.Updates;
```

Add fields near existing `RibbonAnalyticsHelper analytics;`:

```csharp
        private bool isBoundToUpdateNotificationService;
        private Image aboutButtonImage;
        private Image aboutButtonImageWithUpdate;
```

In `AgentRibbon_Load`, after `ApplyLocalizedLabels();`, add:

```csharp
            ApplyAboutButtonImage();
```

Replace `AboutButton_Click` with:

```csharp
        private void AboutButton_Click(object sender, RibbonControlEventArgs e)
        {
            TrackRibbonClick("ribbon.about.clicked");
            var result = AboutDialog.Show(CreateAboutDialogModel(), GetStrings());
            if (result == AboutDialogAction.IgnoreVersion)
            {
                Globals.ThisAddIn?.UpdateNotificationService?.IgnoreCurrentVersion();
                TrackRibbonClick(
                    "ribbon.version_update.ignored",
                    new Dictionary<string, object>(StringComparer.Ordinal)
                    {
                        ["latestVersion"] = Globals.ThisAddIn?.UpdateNotificationService?.CurrentState?.LatestVersion ?? string.Empty,
                    });
            }
        }
```

Add this method after `CreateAboutMessage()`:

```csharp
        private static AboutDialogModel CreateAboutDialogModel()
        {
            var assembly = typeof(AgentRibbon).Assembly;
            var strings = GetStrings();
            var assemblyVersion = assembly.GetName().Version?.ToString() ?? strings.UnknownText;
            var updateState = Globals.ThisAddIn?.UpdateNotificationService?.CurrentState ?? UpdateNotificationState.Empty;

            return new AboutDialogModel
            {
                AppVersion = VersionInfo.AppVersion,
                AssemblyVersion = assemblyVersion,
                BuildConfiguration = GetBuildConfiguration(),
                BuildTime = GetAssemblyBuildTime(assembly),
                HasNewVersion = updateState.HasNewVersion,
                LatestVersion = updateState.LatestVersion,
                DownloadUrl = updateState.DownloadUrl,
                ReleaseNotesUrl = updateState.ReleaseNotesUrl,
                PublishedAtUtc = updateState.PublishedAtUtc,
                UpdateTitle = updateState.Title,
                UpdateSummary = updateState.Summary,
            };
        }
```

In `BindToControllersAndRefresh`, after `EnsureAnalyticsHelper();`, add:

```csharp
            TryBindToUpdateNotificationService();
```

Add these methods near the controller binding helpers:

```csharp
        private bool TryBindToUpdateNotificationService()
        {
            if (isBoundToUpdateNotificationService)
            {
                return true;
            }

            var service = Globals.ThisAddIn?.UpdateNotificationService;
            if (service == null)
            {
                return false;
            }

            service.StateChanged += UpdateNotificationService_StateChanged;
            isBoundToUpdateNotificationService = true;
            ApplyAboutButtonImage();
            return true;
        }

        private void UpdateNotificationService_StateChanged(object sender, EventArgs e)
        {
            ApplyAboutButtonImage();
        }

        private void ApplyAboutButtonImage()
        {
            var hasUpdate = Globals.ThisAddIn?.UpdateNotificationService?.CurrentState?.HasNewVersion == true;
            aboutButton.OfficeImageId = string.Empty;
            aboutButton.Image = hasUpdate ? GetAboutButtonImageWithUpdate() : GetAboutButtonImage();
            aboutButton.ShowImage = true;
            RibbonUI?.InvalidateControl(aboutButton.Name);
        }

        private Image GetAboutButtonImage()
        {
            return aboutButtonImage ?? (aboutButtonImage = RibbonAboutIconFactory.CreateAboutIcon(hasUpdate: false));
        }

        private Image GetAboutButtonImageWithUpdate()
        {
            return aboutButtonImageWithUpdate ?? (aboutButtonImageWithUpdate = RibbonAboutIconFactory.CreateAboutIcon(hasUpdate: true));
        }
```

- [ ] **Step 8: Run Ribbon integration tests**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~AgentRibbonBindsUpdateNotificationServiceAndRefreshesAboutIcon|FullyQualifiedName~AboutButtonOpensUpdateAwareDialogAndCanIgnoreVersion|FullyQualifiedName~HostLocalizedStringsIncludeUpdateReminderText"
```

Expected: PASS.

- [ ] **Step 9: Commit Task 5**

```powershell
git add src/OfficeAgent.ExcelAddIn/RibbonAboutIconFactory.cs src/OfficeAgent.ExcelAddIn/Dialogs/AboutDialog.cs src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj src/OfficeAgent.ExcelAddIn/Localization/HostLocalizedStrings.cs src/OfficeAgent.ExcelAddIn/AgentRibbon.cs tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs
git commit -m "feat: show update reminder on about ribbon"
```

---

### Task 6: Documentation Updates

**Files:**
- Modify: `docs/modules/ribbon-sync-current-behavior.md`
- Modify: `docs/vsto-manual-test-checklist.md`
- Modify: `docs/ribbon-button-custom-icons-guide.md`

- [ ] **Step 1: Update Ribbon current behavior**

In `docs/modules/ribbon-sync-current-behavior.md`, update the Help group section in Section 2 to mention:

```markdown
- 帮助 / `Help`
  - `文档` / `Documentation`
  - `关于` / `About`

Release 安装包会在后台以非阻塞方式检查新版本。检查源是内部配置的独立 URL，不复用业务后端，也不显示在任务窗格设置页。发现比当前安装版本更高的新版本且用户未忽略该版本时，`关于` / `About` 图标会显示红点。点击 `关于` 会展示当前版本、最新版本、下载入口、发布说明入口，并可选择 `忽略此版本` / `Ignore this version`。Debug / 本地开发刷新环境不会请求更新源。更新检查失败只写本地日志，不弹窗，也不影响 Ribbon、任务窗格、登录、同步或模板操作。
```

Also update the icon paragraph to state:

```markdown
`关于` / `About` 使用宿主生成的自定义图标，以便在存在新版本时叠加红点；其他 Ribbon 按钮仍使用 Office 内置 `imageMso` 图标。
```

- [ ] **Step 2: Update manual test checklist**

In `docs/vsto-manual-test-checklist.md`, add a manual case near the Ribbon Help validation:

```markdown
- Release 安装包更新提醒：配置内部更新 manifest URL，使其返回 `Content-Type: application/octet-stream` 的 JSON 字节流，且 `latestVersion` 高于当前 `VersionInfo.AppVersion`。打开 Excel 后确认 `关于` / `About` 图标显示红点；点击 `关于` 后确认显示当前版本、最新版本、下载入口和发布说明入口；点击 `忽略此版本` / `Ignore this version` 后确认红点消失；把 manifest 提高到更高版本后确认红点重新出现。
- 更新检查失败隔离：让更新 manifest URL 断开或返回非法 JSON，重新打开 Excel，确认 Ribbon、任务窗格、登录、下载、上传和模板操作仍可用，且没有更新失败弹窗。
- Debug 环境隔离：运行 `eng/Dev-RefreshExcelAddIn.ps1 -CloseExcel` 后打开 Excel，确认不会请求更新 manifest URL。
```

- [ ] **Step 3: Update custom icon guide**

In `docs/ribbon-button-custom-icons-guide.md`, add a note to the current status or testing section:

```markdown
例外：`关于` / `About` 在新版本提醒功能中会由运行时代码设置为宿主生成的自定义图标。无更新时显示普通信息图标；有未忽略的新版本时显示带红点的信息图标。该例外只改变 `aboutButton` 图片来源，不改变按钮标签、本地化、点击行为或大按钮布局。
```

- [ ] **Step 4: Review docs diff**

Run:

```powershell
git diff -- docs/modules/ribbon-sync-current-behavior.md docs/vsto-manual-test-checklist.md docs/ribbon-button-custom-icons-guide.md
```

Expected: diff only describes the new update reminder behavior and the About icon exception.

- [ ] **Step 5: Commit Task 6**

```powershell
git add docs/modules/ribbon-sync-current-behavior.md docs/vsto-manual-test-checklist.md docs/ribbon-button-custom-icons-guide.md
git commit -m "docs: document ribbon version reminder"
```

---

### Task 7: Full Verification

**Files:**
- No code changes expected.

- [ ] **Step 1: Run ExcelAddIn test suite**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj
```

Expected: PASS.

- [ ] **Step 2: Run Debug VSTO refresh build**

Run:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File eng/Dev-RefreshExcelAddIn.ps1 -CloseExcel
```

Expected: PASS. This validates the Debug add-in still builds and refreshes local registration. Debug should not request the update manifest.

- [ ] **Step 3: Inspect final status**

Run:

```powershell
git status --short
git diff --stat
```

Expected: clean working tree after the previous task commits. If verification writes generated output, inspect it and commit only intentional source, test, or documentation changes.

---

## Self-Review

Spec coverage:

- Release-only gate: Task 3 adds `UpdateCheckConfiguration`; Task 4 composes it; Task 7 validates Debug refresh.
- Non-blocking behavior: Task 3 tests `StartBackgroundCheck` returns before delayed request; Task 4 asserts `ThisAddIn` does not await update checks.
- Independent URL and octet-stream JSON: Task 2 validates JSON parsing with `application/octet-stream`; Task 3 uses `UpdateCheckOptions.ManifestUrl`.
- Cache and ignore state: Task 3 covers 24-hour cache, ignored version, and higher future version.
- About red dot: Task 5 binds update state to `aboutButton` and runtime-generated red-dot image.
- User-visible About flow: Task 5 adds `AboutDialog` with download, release notes, and ignore action.
- Docs and manual validation: Task 6 updates current behavior, manual checklist, and custom icon guide.

Placeholder scan:

- The plan contains no `TBD`, no `TODO`, no unnamed files, and no unspecified test commands.
- The default manifest URL is intentionally `string.Empty`; Release checks run when the installer or deployment writes `HKCU\Software\OfficeAgent\UpdateManifestUrl` or `HKLM\Software\OfficeAgent\UpdateManifestUrl`.

Type consistency:

- `UpdateManifest.LatestVersion`, `UpdateState.LatestVersion`, and `UpdateNotificationState.LatestVersion` use the same property name.
- `UpdateNotificationService.CurrentState` is the single state read by `AgentRibbon`.
- `IgnoreCurrentVersion()` is the only ignore write path used by `AboutButton_Click`.
