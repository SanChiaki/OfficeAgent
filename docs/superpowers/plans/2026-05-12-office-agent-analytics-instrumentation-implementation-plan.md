# OfficeAgent Analytics Instrumentation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add HTTP-backed analytics instrumentation for Ribbon, task pane panel, bridge, and connector flows using the `/insertLog` envelope described in the approved design.

**Architecture:** Add a separate Core analytics abstraction, an Infrastructure HTTP sink, and host/frontend call sites that never block user workflows. The outer `/insertLog` payload is produced only in Infrastructure; Ribbon, Panel, Bridge, and Connector code emit structured `AnalyticsEvent` data with stable `properties` and optional `businessContext`.

**Tech Stack:** C# / .NET Framework 4.8, Newtonsoft.Json, VSTO Excel add-in, React + TypeScript + Vite/Vitest, Express mock server.

---

## File Structure

Create:

- `src/OfficeAgent.Core/Analytics/AnalyticsError.cs`  
  Small error DTO for failed analytics events.
- `src/OfficeAgent.Core/Analytics/AnalyticsEvent.cs`  
  Canonical event model serialized into `/insertLog.answer`.
- `src/OfficeAgent.Core/Analytics/IAnalyticsSink.cs`  
  Low-level async writer contract.
- `src/OfficeAgent.Core/Analytics/IAnalyticsService.cs`  
  High-level fire-and-forget tracking API used by app code.
- `src/OfficeAgent.Core/Analytics/AnalyticsService.cs`  
  Adds defaults, invokes sink, catches failures.
- `src/OfficeAgent.Core/Analytics/NoopAnalyticsService.cs`  
  Disabled/default analytics implementation.
- `src/OfficeAgent.Infrastructure/Analytics/InsertLogAnalyticsSink.cs`  
  Sends events to `{AnalyticsBaseUrl}/insertLog`.
- `src/OfficeAgent.ExcelAddIn/Analytics/RibbonAnalyticsHelper.cs`  
  Builds common Ribbon properties and safely emits events.
- `src/OfficeAgent.Frontend/src/analytics/panelAnalytics.ts`  
  Thin panel helper around `nativeBridge.trackAnalytics`.
- `tests/OfficeAgent.Core.Tests/AnalyticsServiceTests.cs`
- `tests/OfficeAgent.Infrastructure.Tests/InsertLogAnalyticsSinkTests.cs`

Modify:

- `src/OfficeAgent.Core/Models/AppSettings.cs`  
  Add `AnalyticsBaseUrl`.
- `src/OfficeAgent.Infrastructure/Storage/FileSettingsStore.cs`  
  Persist and normalize `AnalyticsBaseUrl`.
- `src/OfficeAgent.ExcelAddIn/WebBridge/WebMessageEnvelope.cs`  
  Add `bridge.trackAnalytics` constant and payload DTO.
- `src/OfficeAgent.ExcelAddIn/WebBridge/WebMessageRouter.cs`  
  Route panel analytics payloads to `IAnalyticsService`.
- `src/OfficeAgent.ExcelAddIn/WebBridge/WebViewBootstrapper.cs`  
  Pass analytics service into router.
- `src/OfficeAgent.ExcelAddIn/TaskPane/TaskPaneHostControl.cs`  
  Accept analytics service.
- `src/OfficeAgent.ExcelAddIn/TaskPane/TaskPaneController.cs`  
  Accept analytics service.
- `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs`  
  Compose analytics service and pass it to consumers.
- `src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj`  
  Include new `Analytics\RibbonAnalyticsHelper.cs`.
- `src/OfficeAgent.ExcelAddIn/AgentRibbon.cs`  
  Emit Ribbon entry click/dropdown/login/help/about events.
- `src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs`  
  Emit project, initialize, AI mapping, download, upload result events.
- `src/OfficeAgent.ExcelAddIn/RibbonTemplateController.cs`  
  Emit template command events.
- `src/OfficeAgent.Core/Sync/WorksheetSyncService.cs`  
  Emit connector-level request/completion/failure events.
- `src/OfficeAgent.Infrastructure/Http/CurrentBusinessSystemConnector.cs`  
  Optionally emit connector-specific `business.current.*` events.
- `src/OfficeAgent.Frontend/src/types/bridge.ts`  
  Add analytics payload types and `analyticsBaseUrl` setting.
- `src/OfficeAgent.Frontend/src/bridge/nativeBridge.ts`  
  Add `trackAnalytics` and browser-preview no-op.
- `src/OfficeAgent.Frontend/src/App.tsx`  
  Add settings field and panel event calls.
- `src/OfficeAgent.Frontend/src/i18n/uiStrings.ts`  
  Add localized `Analytics Base URL` label.
- `tests/mock-server/server.js`  
  Add `/insertLog`, `/analytics/logs`, and delete endpoint.
- `tests/mock-server/README.md`
- `docs/modules/task-pane-current-behavior.md`
- `docs/modules/ribbon-sync-current-behavior.md`
- `docs/module-index.md`
- `docs/vsto-manual-test-checklist.md`

Test:

- `tests/OfficeAgent.Core.Tests/AnalyticsServiceTests.cs`
- `tests/OfficeAgent.Infrastructure.Tests/InsertLogAnalyticsSinkTests.cs`
- `tests/OfficeAgent.Infrastructure.Tests/FileSettingsStoreTests.cs`
- `tests/OfficeAgent.ExcelAddIn.Tests/WebMessageRouterTests.cs`
- `tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs`
- `tests/OfficeAgent.ExcelAddIn.Tests/RibbonSyncControllerTests.cs`
- `tests/OfficeAgent.Core.Tests/WorksheetSyncServiceTests.cs`
- `src/OfficeAgent.Frontend/src/bridge/nativeBridge.test.ts`
- `src/OfficeAgent.Frontend/src/App.test.tsx`

---

### Task 1: Core Analytics Contracts And Service

**Files:**
- Create: `src/OfficeAgent.Core/Analytics/AnalyticsError.cs`
- Create: `src/OfficeAgent.Core/Analytics/AnalyticsEvent.cs`
- Create: `src/OfficeAgent.Core/Analytics/IAnalyticsSink.cs`
- Create: `src/OfficeAgent.Core/Analytics/IAnalyticsService.cs`
- Create: `src/OfficeAgent.Core/Analytics/AnalyticsService.cs`
- Create: `src/OfficeAgent.Core/Analytics/NoopAnalyticsService.cs`
- Test: `tests/OfficeAgent.Core.Tests/AnalyticsServiceTests.cs`

- [ ] **Step 1: Write failing Core analytics tests**

Create `tests/OfficeAgent.Core.Tests/AnalyticsServiceTests.cs`:

```csharp
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using OfficeAgent.Core.Analytics;
using OfficeAgent.Core.Diagnostics;
using Xunit;

namespace OfficeAgent.Core.Tests
{
    public sealed class AnalyticsServiceTests : IDisposable
    {
        [Fact]
        public void TrackAddsSchemaVersionAndTimestampBeforeWritingToSink()
        {
            var sink = new RecordingAnalyticsSink();
            var service = new AnalyticsService(sink);

            service.Track("ribbon.download.clicked", "ribbon", new Dictionary<string, object>
            {
                ["projectId"] = "performance",
                ["projectName"] = "绩效项目",
            });

            Assert.True(sink.Written.Wait(TimeSpan.FromSeconds(2)));
            Assert.NotNull(sink.LastEvent);
            Assert.Equal(1, sink.LastEvent.SchemaVersion);
            Assert.Equal("ribbon.download.clicked", sink.LastEvent.EventName);
            Assert.Equal("ribbon", sink.LastEvent.Source);
            Assert.True(DateTime.UtcNow.Subtract(sink.LastEvent.OccurredAtUtc).TotalSeconds < 10);
            Assert.Equal("performance", sink.LastEvent.Properties["projectId"]);
            Assert.Equal("绩效项目", sink.LastEvent.Properties["projectName"]);
        }

        [Fact]
        public void TrackDoesNotThrowWhenSinkFailsAndWritesDiagnosticLog()
        {
            var entries = new List<OfficeAgentLogEntry>();
            OfficeAgentLog.Configure(entries.Add);
            var service = new AnalyticsService(new FailingAnalyticsSink());

            service.Track("panel.settings.saved", "panel");

            Assert.True(SpinWait.SpinUntil(
                () => entries.Exists(entry => entry.Component == "analytics" && entry.EventName == "track.failed"),
                TimeSpan.FromSeconds(2)));
            Assert.Contains(entries, entry =>
                entry.Component == "analytics" &&
                entry.EventName == "track.failed" &&
                entry.Level == "warn");
        }

        [Fact]
        public void NoopAnalyticsServiceAcceptsEventsWithoutWriting()
        {
            NoopAnalyticsService.Instance.Track("ribbon.about.clicked", "ribbon");
            NoopAnalyticsService.Instance.Track(new AnalyticsEvent { EventName = "panel.opened", Source = "panel" });
        }

        public void Dispose()
        {
            OfficeAgentLog.Reset();
        }

        private sealed class RecordingAnalyticsSink : IAnalyticsSink
        {
            public ManualResetEventSlim Written { get; } = new ManualResetEventSlim();

            public AnalyticsEvent LastEvent { get; private set; }

            public Task WriteAsync(AnalyticsEvent analyticsEvent, CancellationToken cancellationToken)
            {
                LastEvent = analyticsEvent;
                Written.Set();
                return Task.CompletedTask;
            }
        }

        private sealed class FailingAnalyticsSink : IAnalyticsSink
        {
            public Task WriteAsync(AnalyticsEvent analyticsEvent, CancellationToken cancellationToken)
            {
                throw new InvalidOperationException("sink failed");
            }
        }
    }
}
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```powershell
dotnet test tests/OfficeAgent.Core.Tests/OfficeAgent.Core.Tests.csproj --filter FullyQualifiedName~AnalyticsServiceTests
```

Expected: FAIL because `OfficeAgent.Core.Analytics` types do not exist.

- [ ] **Step 3: Add Core analytics implementation**

Create `src/OfficeAgent.Core/Analytics/AnalyticsError.cs`:

```csharp
namespace OfficeAgent.Core.Analytics
{
    public sealed class AnalyticsError
    {
        public string Code { get; set; } = string.Empty;

        public string Message { get; set; } = string.Empty;

        public string ExceptionType { get; set; } = string.Empty;
    }
}
```

Create `src/OfficeAgent.Core/Analytics/AnalyticsEvent.cs`:

```csharp
using System;
using System.Collections.Generic;

namespace OfficeAgent.Core.Analytics
{
    public sealed class AnalyticsEvent
    {
        public int SchemaVersion { get; set; } = 1;

        public string EventName { get; set; } = string.Empty;

        public string Source { get; set; } = string.Empty;

        public DateTime OccurredAtUtc { get; set; }

        public IDictionary<string, object> Properties { get; set; } = new Dictionary<string, object>(StringComparer.Ordinal);

        public IDictionary<string, object> BusinessContext { get; set; } = new Dictionary<string, object>(StringComparer.Ordinal);

        public AnalyticsError Error { get; set; }
    }
}
```

Create `src/OfficeAgent.Core/Analytics/IAnalyticsSink.cs`:

```csharp
using System.Threading;
using System.Threading.Tasks;

namespace OfficeAgent.Core.Analytics
{
    public interface IAnalyticsSink
    {
        Task WriteAsync(AnalyticsEvent analyticsEvent, CancellationToken cancellationToken);
    }
}
```

Create `src/OfficeAgent.Core/Analytics/IAnalyticsService.cs`:

```csharp
using System.Collections.Generic;

namespace OfficeAgent.Core.Analytics
{
    public interface IAnalyticsService
    {
        void Track(AnalyticsEvent analyticsEvent);

        void Track(
            string eventName,
            string source,
            IDictionary<string, object> properties = null,
            IDictionary<string, object> businessContext = null,
            AnalyticsError error = null);
    }
}
```

Create `src/OfficeAgent.Core/Analytics/AnalyticsService.cs`:

```csharp
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using OfficeAgent.Core.Diagnostics;

namespace OfficeAgent.Core.Analytics
{
    public sealed class AnalyticsService : IAnalyticsService
    {
        private readonly IAnalyticsSink sink;

        public AnalyticsService(IAnalyticsSink sink)
        {
            this.sink = sink ?? throw new ArgumentNullException(nameof(sink));
        }

        public void Track(AnalyticsEvent analyticsEvent)
        {
            if (analyticsEvent == null || string.IsNullOrWhiteSpace(analyticsEvent.EventName))
            {
                return;
            }

            var normalized = Normalize(analyticsEvent);
            Task.Run(async () =>
            {
                try
                {
                    await sink.WriteAsync(normalized, CancellationToken.None).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    OfficeAgentLog.Warn("analytics", "track.failed", $"Analytics event failed: {normalized.EventName}", ex.Message);
                }
            });
        }

        public void Track(
            string eventName,
            string source,
            IDictionary<string, object> properties = null,
            IDictionary<string, object> businessContext = null,
            AnalyticsError error = null)
        {
            Track(new AnalyticsEvent
            {
                EventName = eventName ?? string.Empty,
                Source = source ?? string.Empty,
                Properties = properties ?? new Dictionary<string, object>(StringComparer.Ordinal),
                BusinessContext = businessContext ?? new Dictionary<string, object>(StringComparer.Ordinal),
                Error = error,
            });
        }

        private static AnalyticsEvent Normalize(AnalyticsEvent analyticsEvent)
        {
            analyticsEvent.SchemaVersion = analyticsEvent.SchemaVersion <= 0 ? 1 : analyticsEvent.SchemaVersion;
            analyticsEvent.Source = analyticsEvent.Source ?? string.Empty;
            analyticsEvent.Properties = analyticsEvent.Properties ?? new Dictionary<string, object>(StringComparer.Ordinal);
            analyticsEvent.BusinessContext = analyticsEvent.BusinessContext ?? new Dictionary<string, object>(StringComparer.Ordinal);
            if (analyticsEvent.OccurredAtUtc == default)
            {
                analyticsEvent.OccurredAtUtc = DateTime.UtcNow;
            }

            return analyticsEvent;
        }
    }
}
```

Create `src/OfficeAgent.Core/Analytics/NoopAnalyticsService.cs`:

```csharp
using System.Collections.Generic;

namespace OfficeAgent.Core.Analytics
{
    public sealed class NoopAnalyticsService : IAnalyticsService
    {
        public static readonly NoopAnalyticsService Instance = new NoopAnalyticsService();

        private NoopAnalyticsService()
        {
        }

        public void Track(AnalyticsEvent analyticsEvent)
        {
        }

        public void Track(
            string eventName,
            string source,
            IDictionary<string, object> properties = null,
            IDictionary<string, object> businessContext = null,
            AnalyticsError error = null)
        {
        }
    }
}
```

- [ ] **Step 4: Run Core analytics tests**

Run:

```powershell
dotnet test tests/OfficeAgent.Core.Tests/OfficeAgent.Core.Tests.csproj --filter FullyQualifiedName~AnalyticsServiceTests
```

Expected: PASS.

- [ ] **Step 5: Commit Core analytics contracts**

```powershell
git add src/OfficeAgent.Core/Analytics tests/OfficeAgent.Core.Tests/AnalyticsServiceTests.cs
git commit -m "feat: add analytics core service"
```

---

### Task 2: Settings Persistence For Analytics Base URL

**Files:**
- Modify: `src/OfficeAgent.Core/Models/AppSettings.cs`
- Modify: `src/OfficeAgent.Infrastructure/Storage/FileSettingsStore.cs`
- Test: `tests/OfficeAgent.Infrastructure.Tests/FileSettingsStoreTests.cs`

- [ ] **Step 1: Add failing settings tests**

Append to `tests/OfficeAgent.Infrastructure.Tests/FileSettingsStoreTests.cs`:

```csharp
[Fact]
public void SaveRoundTripsAnalyticsBaseUrl()
{
    var settingsPath = Path.Combine(tempDirectory, "settings.json");
    var store = new FileSettingsStore(settingsPath, new DpapiSecretProtector());

    store.Save(new AppSettings
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
```

- [ ] **Step 2: Run settings tests to verify they fail**

Run:

```powershell
dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj --filter "FullyQualifiedName~SaveRoundTripsAnalyticsBaseUrl|FullyQualifiedName~LoadDefaultsAnalyticsBaseUrlToEmptyString"
```

Expected: FAIL because `AppSettings.AnalyticsBaseUrl` does not exist.

- [ ] **Step 3: Add AnalyticsBaseUrl to AppSettings**

Modify `src/OfficeAgent.Core/Models/AppSettings.cs` by adding the property after `BusinessBaseUrl`:

```csharp
public string AnalyticsBaseUrl { get; set; } = string.Empty;
```

Use the existing `NormalizeOptionalUrl` method for this field.

- [ ] **Step 4: Persist AnalyticsBaseUrl**

Modify `src/OfficeAgent.Infrastructure/Storage/FileSettingsStore.cs`.

In `Load()`, add:

```csharp
AnalyticsBaseUrl = AppSettings.NormalizeOptionalUrl(persisted.AnalyticsBaseUrl),
```

In `Save(AppSettings settings)`, add:

```csharp
AnalyticsBaseUrl = AppSettings.NormalizeOptionalUrl(settings?.AnalyticsBaseUrl),
```

In `PersistedSettings`, add:

```csharp
public string AnalyticsBaseUrl { get; set; } = string.Empty;
```

- [ ] **Step 5: Run settings tests**

Run:

```powershell
dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj --filter "FullyQualifiedName~SaveRoundTripsAnalyticsBaseUrl|FullyQualifiedName~LoadDefaultsAnalyticsBaseUrlToEmptyString"
```

Expected: PASS.

- [ ] **Step 6: Commit settings persistence**

```powershell
git add src/OfficeAgent.Core/Models/AppSettings.cs src/OfficeAgent.Infrastructure/Storage/FileSettingsStore.cs tests/OfficeAgent.Infrastructure.Tests/FileSettingsStoreTests.cs
git commit -m "feat: persist analytics endpoint setting"
```

---

### Task 3: InsertLog HTTP Analytics Sink

**Files:**
- Create: `src/OfficeAgent.Infrastructure/Analytics/InsertLogAnalyticsSink.cs`
- Test: `tests/OfficeAgent.Infrastructure.Tests/InsertLogAnalyticsSinkTests.cs`

- [ ] **Step 1: Write failing sink tests**

Create `tests/OfficeAgent.Infrastructure.Tests/InsertLogAnalyticsSinkTests.cs`:

```csharp
using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeAgent.Core.Analytics;
using OfficeAgent.Core.Models;
using OfficeAgent.Infrastructure.Analytics;
using Xunit;

namespace OfficeAgent.Infrastructure.Tests
{
    public sealed class InsertLogAnalyticsSinkTests
    {
        [Fact]
        public async Task WriteAsyncPostsInsertLogEnvelopeWithJsonAnswer()
        {
            var handler = new RecordingHandler(_ => new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent("{\"ok\":true}"),
            });
            var sink = new InsertLogAnalyticsSink(
                () => new AppSettings { AnalyticsBaseUrl = "https://analytics.internal.example/v1/" },
                new HttpClient(handler));

            await sink.WriteAsync(new AnalyticsEvent
            {
                EventName = "ribbon.download.clicked",
                Source = "ribbon",
                Properties = new Dictionary<string, object>
                {
                    ["projectId"] = "performance",
                    ["projectName"] = "绩效项目",
                },
            }, CancellationToken.None);

            Assert.Equal("https://analytics.internal.example/v1/insertLog", handler.LastRequest.RequestUri.ToString());
            var body = JObject.Parse(handler.LastBody);
            Assert.Equal("excelAi", (string)body["frontEndIntent"]);
            Assert.Equal("Excel", (string)body["clientSource"]);
            Assert.Equal(1, (int)body["questionType"]);
            Assert.False(string.IsNullOrWhiteSpace((string)body["askId"]));
            Assert.False(string.IsNullOrWhiteSpace((string)body["talkId"]));

            var answer = JObject.Parse((string)body["answer"]);
            Assert.Equal("ribbon.download.clicked", (string)answer["eventName"]);
            Assert.Equal("绩效项目", (string)answer["properties"]["projectName"]);
        }

        [Fact]
        public async Task WriteAsyncRejectsMissingAnalyticsBaseUrl()
        {
            var handler = new RecordingHandler(_ => new HttpResponseMessage(HttpStatusCode.OK));
            var sink = new InsertLogAnalyticsSink(
                () => new AppSettings { AnalyticsBaseUrl = " " },
                new HttpClient(handler));

            var error = await Assert.ThrowsAsync<InvalidOperationException>(() =>
                sink.WriteAsync(new AnalyticsEvent { EventName = "panel.opened", Source = "panel" }, CancellationToken.None));

            Assert.Equal("The configured Analytics Base URL is invalid. Update settings and try again.", error.Message);
            Assert.Equal(0, handler.CallCount);
        }

        [Fact]
        public async Task WriteAsyncThrowsForNonSuccessResponse()
        {
            var handler = new RecordingHandler(_ => new HttpResponseMessage(HttpStatusCode.BadRequest)
            {
                Content = new StringContent("bad request"),
            });
            var sink = new InsertLogAnalyticsSink(
                () => new AppSettings { AnalyticsBaseUrl = "https://analytics.internal.example" },
                new HttpClient(handler));

            var error = await Assert.ThrowsAsync<InvalidOperationException>(() =>
                sink.WriteAsync(new AnalyticsEvent { EventName = "panel.opened", Source = "panel" }, CancellationToken.None));

            Assert.Contains("Analytics request failed (400 Bad Request): bad request", error.Message, StringComparison.Ordinal);
        }

        private sealed class RecordingHandler : HttpMessageHandler
        {
            private readonly Func<HttpRequestMessage, HttpResponseMessage> responder;

            public RecordingHandler(Func<HttpRequestMessage, HttpResponseMessage> responder)
            {
                this.responder = responder;
            }

            public HttpRequestMessage LastRequest { get; private set; }

            public string LastBody { get; private set; } = string.Empty;

            public int CallCount { get; private set; }

            protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
            {
                CallCount++;
                LastRequest = request;
                LastBody = request.Content == null ? string.Empty : await request.Content.ReadAsStringAsync();
                return responder(request);
            }
        }
    }
}
```

- [ ] **Step 2: Run sink tests to verify they fail**

Run:

```powershell
dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj --filter FullyQualifiedName~InsertLogAnalyticsSinkTests
```

Expected: FAIL because `OfficeAgent.Infrastructure.Analytics.InsertLogAnalyticsSink` does not exist.

- [ ] **Step 3: Implement InsertLogAnalyticsSink**

Create `src/OfficeAgent.Infrastructure/Analytics/InsertLogAnalyticsSink.cs`:

```csharp
using System;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using OfficeAgent.Core.Analytics;
using OfficeAgent.Core.Models;

namespace OfficeAgent.Infrastructure.Analytics
{
    public sealed class InsertLogAnalyticsSink : IAnalyticsSink
    {
        private static readonly JsonSerializerSettings SerializerSettings = new JsonSerializerSettings
        {
            ContractResolver = new CamelCasePropertyNamesContractResolver(),
            NullValueHandling = NullValueHandling.Ignore,
        };

        private readonly Func<AppSettings> loadSettings;
        private readonly HttpClient httpClient;

        public InsertLogAnalyticsSink(Func<AppSettings> loadSettings, HttpClient httpClient = null)
        {
            this.loadSettings = loadSettings ?? throw new ArgumentNullException(nameof(loadSettings));
            this.httpClient = httpClient ?? new HttpClient { Timeout = TimeSpan.FromSeconds(5) };
        }

        public async Task WriteAsync(AnalyticsEvent analyticsEvent, CancellationToken cancellationToken)
        {
            if (analyticsEvent == null)
            {
                throw new ArgumentNullException(nameof(analyticsEvent));
            }

            var settings = loadSettings() ?? new AppSettings();
            var baseUrl = AppSettings.NormalizeOptionalUrl(settings.AnalyticsBaseUrl);
            if (!Uri.TryCreate(baseUrl, UriKind.Absolute, out var baseUri) ||
                (baseUri.Scheme != Uri.UriSchemeHttp && baseUri.Scheme != Uri.UriSchemeHttps))
            {
                throw new InvalidOperationException("The configured Analytics Base URL is invalid. Update settings and try again.");
            }

            var endpoint = new Uri($"{baseUri.AbsoluteUri.TrimEnd('/')}/insertLog");
            var envelope = new
            {
                frontEndIntent = "excelAi",
                clientSource = "Excel",
                questionType = 1,
                askId = CreateId(),
                talkId = CreateId(),
                answer = JsonConvert.SerializeObject(analyticsEvent, SerializerSettings),
            };

            var payload = JsonConvert.SerializeObject(envelope, SerializerSettings);
            using (var request = new HttpRequestMessage(HttpMethod.Post, endpoint))
            {
                request.Content = new StringContent(payload, Encoding.UTF8, "application/json");
                using (var response = await httpClient.SendAsync(request, cancellationToken).ConfigureAwait(false))
                {
                    if (response.IsSuccessStatusCode)
                    {
                        return;
                    }

                    var responseBody = response.Content == null
                        ? string.Empty
                        : await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    throw new InvalidOperationException(
                        $"Analytics request failed ({(int)response.StatusCode} {response.ReasonPhrase}): {responseBody}");
                }
            }
        }

        private static string CreateId()
        {
            var bytes = new byte[24];
            using (var generator = RandomNumberGenerator.Create())
            {
                generator.GetBytes(bytes);
            }

            return Convert.ToBase64String(bytes)
                .TrimEnd('=')
                .Replace('+', '-')
                .Replace('/', '_');
        }
    }
}
```

- [ ] **Step 4: Run sink tests**

Run:

```powershell
dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj --filter FullyQualifiedName~InsertLogAnalyticsSinkTests
```

Expected: PASS.

- [ ] **Step 5: Commit HTTP sink**

```powershell
git add src/OfficeAgent.Infrastructure/Analytics/InsertLogAnalyticsSink.cs tests/OfficeAgent.Infrastructure.Tests/InsertLogAnalyticsSinkTests.cs
git commit -m "feat: add insertLog analytics sink"
```

---

### Task 4: Bridge Analytics Message

**Files:**
- Modify: `src/OfficeAgent.ExcelAddIn/WebBridge/WebMessageEnvelope.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/WebBridge/WebMessageRouter.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/WebBridge/WebViewBootstrapper.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/TaskPane/TaskPaneHostControl.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/TaskPane/TaskPaneController.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs`
- Test: `tests/OfficeAgent.ExcelAddIn.Tests/WebMessageRouterTests.cs`

- [ ] **Step 1: Add failing WebMessageRouter analytics tests**

Add `using OfficeAgent.Core.Analytics;` to `tests/OfficeAgent.ExcelAddIn.Tests/WebMessageRouterTests.cs`.

Append tests:

```csharp
[Fact]
public void TrackAnalyticsRoutesPanelEventToAnalyticsService()
{
    var sessionStore = new FileSessionStore(Path.Combine(tempDirectory, "sessions"));
    var settingsStore = new FileSettingsStore(
        Path.Combine(tempDirectory, "settings.json"),
        new DpapiSecretProtector());
    var analytics = new RecordingAnalyticsService();
    var router = CreateRouter(sessionStore, settingsStore, analyticsService: analytics, resolvedUiLocale: "zh");

    var responseJson = InvokeRoute(
        router,
        "{\"type\":\"bridge.trackAnalytics\",\"requestId\":\"req-1\",\"payload\":{\"eventName\":\"panel.composer.send.clicked\",\"source\":\"panel\",\"properties\":{\"inputLength\":12},\"businessContext\":{\"module\":\"demo\"}}}");

    Assert.Contains("\"ok\":true", responseJson);
    Assert.Equal("panel.composer.send.clicked", analytics.LastEvent.EventName);
    Assert.Equal("panel", analytics.LastEvent.Source);
    Assert.Equal(12L, analytics.LastEvent.Properties["inputLength"]);
    Assert.Equal("zh", analytics.LastEvent.Properties["uiLocale"]);
    Assert.Equal("demo", analytics.LastEvent.BusinessContext["module"]);
}

[Fact]
public void TrackAnalyticsRejectsBlankEventName()
{
    var sessionStore = new FileSessionStore(Path.Combine(tempDirectory, "sessions"));
    var settingsStore = new FileSettingsStore(
        Path.Combine(tempDirectory, "settings.json"),
        new DpapiSecretProtector());
    var analytics = new RecordingAnalyticsService();
    var router = CreateRouter(sessionStore, settingsStore, analyticsService: analytics);

    var responseJson = InvokeRoute(
        router,
        "{\"type\":\"bridge.trackAnalytics\",\"requestId\":\"req-1\",\"payload\":{\"eventName\":\" \",\"source\":\"panel\"}}");

    Assert.Contains("\"ok\":false", responseJson);
    Assert.Contains("\"code\":\"malformed_payload\"", responseJson);
    Assert.Null(analytics.LastEvent);
}
```

Add the fake service at the bottom of the test class:

```csharp
private sealed class RecordingAnalyticsService : IAnalyticsService
{
    public AnalyticsEvent LastEvent { get; private set; }

    public void Track(AnalyticsEvent analyticsEvent)
    {
        LastEvent = analyticsEvent;
    }

    public void Track(
        string eventName,
        string source,
        IDictionary<string, object> properties = null,
        IDictionary<string, object> businessContext = null,
        AnalyticsError error = null)
    {
        LastEvent = new AnalyticsEvent
        {
            EventName = eventName,
            Source = source,
            Properties = properties,
            BusinessContext = businessContext,
            Error = error,
        };
    }
}
```

Update the final `CreateRouter(...)` helper overload in `tests/OfficeAgent.ExcelAddIn.Tests/WebMessageRouterTests.cs` so its signature ends with:

```csharp
Func<AppSettings, string> getResolvedUiLocale,
IAnalyticsService analyticsService = null)
```

Then pass this final argument to `Activator.CreateInstance`:

```csharp
args: new object[]
{
    sessionStore,
    settingsStore,
    selectionContextService,
    excelCommandExecutor,
    agentOrchestrator,
    sharedCookies,
    cookieStore,
    getResolvedUiLocale,
    analyticsService,
},
```

Add this overload used by the new tests:

```csharp
private static object CreateRouter(
    FileSessionStore sessionStore,
    FileSettingsStore settingsStore,
    IAnalyticsService analyticsService,
    string resolvedUiLocale = "en")
{
    return CreateRouter(
        sessionStore,
        settingsStore,
        new FakeExcelContextService(SelectionContext.Empty("No selection available.")),
        new FakeExcelCommandExecutor(),
        new FakeAgentOrchestrator(),
        settings => resolvedUiLocale,
        analyticsService);
}
```

- [ ] **Step 2: Run router tests to verify they fail**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~TrackAnalyticsRoutesPanelEventToAnalyticsService|FullyQualifiedName~TrackAnalyticsRejectsBlankEventName"
```

Expected: FAIL because `bridge.trackAnalytics` is unknown and the router constructor does not accept analytics.

- [ ] **Step 3: Add bridge message type and payload**

Modify `src/OfficeAgent.ExcelAddIn/WebBridge/WebMessageEnvelope.cs`.

Add:

```csharp
public const string TrackAnalytics = "bridge.trackAnalytics";
```

Add payload DTO:

```csharp
internal sealed class AnalyticsPayload
{
    public string EventName { get; set; } = string.Empty;

    public string Source { get; set; } = "panel";

    public IDictionary<string, object> Properties { get; set; } = new Dictionary<string, object>();

    public IDictionary<string, object> BusinessContext { get; set; } = new Dictionary<string, object>();
}
```

- [ ] **Step 4: Route analytics in WebMessageRouter**

Modify `src/OfficeAgent.ExcelAddIn/WebBridge/WebMessageRouter.cs`:

- Add `using OfficeAgent.Core.Analytics;`
- Add field:

```csharp
private readonly IAnalyticsService analyticsService;
```

- Add `BridgeMessageTypes.TrackAnalytics` to `allowedTypes`.
- Add optional constructor parameter:

```csharp
IAnalyticsService analyticsService = null
```

- Assign:

```csharp
this.analyticsService = analyticsService ?? NoopAnalyticsService.Instance;
```

- Add switch case:

```csharp
case BridgeMessageTypes.TrackAnalytics:
    return TrackAnalytics(request);
```

- Add method:

```csharp
private WebMessageResponse TrackAnalytics(WebMessageRequest request)
{
    if (request.Payload == null || request.Payload.Type != JTokenType.Object || !request.Payload.HasValues)
    {
        return Error(request.Type, request.RequestId, "malformed_payload", GetStrings().BridgePayloadRequiredMessage(BridgeMessageTypes.TrackAnalytics, "an analytics payload"));
    }

    try
    {
        var payload = request.Payload.ToObject<AnalyticsPayload>() ?? new AnalyticsPayload();
        if (string.IsNullOrWhiteSpace(payload.EventName))
        {
            return Error(request.Type, request.RequestId, "malformed_payload", GetStrings().BridgeValidPayloadRequiredMessage(BridgeMessageTypes.TrackAnalytics, "an analytics payload"));
        }

        var properties = payload.Properties ?? new Dictionary<string, object>();
        properties["uiLocale"] = GetStrings().Locale;
        analyticsService.Track(payload.EventName, string.IsNullOrWhiteSpace(payload.Source) ? "panel" : payload.Source, properties, payload.BusinessContext);
        return Success(request.Type, request.RequestId, new { tracked = true });
    }
    catch (JsonException)
    {
        return Error(request.Type, request.RequestId, "malformed_payload", GetStrings().BridgeValidPayloadRequiredMessage(BridgeMessageTypes.TrackAnalytics, "an analytics payload"));
    }
}
```

`HostLocalizedStrings.Locale` already exists, so use `GetStrings().Locale` for the bridge-added `uiLocale` property.

- [ ] **Step 5: Pass analytics through task pane host classes**

Modify constructors in:

- `src/OfficeAgent.ExcelAddIn/WebBridge/WebViewBootstrapper.cs`
- `src/OfficeAgent.ExcelAddIn/TaskPane/TaskPaneHostControl.cs`
- `src/OfficeAgent.ExcelAddIn/TaskPane/TaskPaneController.cs`

Add `using OfficeAgent.Core.Analytics;` to each file.

In `WebViewBootstrapper`, change the constructor signature to:

```csharp
public WebViewBootstrapper(
    WebView2 webView,
    FileSessionStore sessionStore,
    FileSettingsStore settingsStore,
    IExcelContextService excelContextService,
    IExcelCommandExecutor excelCommandExecutor,
    IAgentOrchestrator agentOrchestrator,
    SharedCookieContainer sharedCookies,
    FileCookieStore cookieStore,
    Func<AppSettings, string> getResolvedUiLocale,
    IAnalyticsService analyticsService = null)
```

and pass the final argument into `WebMessageRouter`:

```csharp
messageRouter = new WebMessageRouter(
    sessionStore,
    settingsStore,
    excelContextService,
    excelCommandExecutor,
    agentOrchestrator,
    sharedCookies,
    cookieStore,
    getResolvedUiLocale,
    analyticsService);
```

In `TaskPaneHostControl`, add `IAnalyticsService analyticsService = null` as the final constructor argument and pass it into `new WebViewBootstrapper(...)`.

In `TaskPaneController`, add a private field:

```csharp
private readonly IAnalyticsService analyticsService;
```

Add `IAnalyticsService analyticsService = null` as the final constructor argument, assign:

```csharp
this.analyticsService = analyticsService ?? NoopAnalyticsService.Instance;
```

and pass `this.analyticsService` into `new TaskPaneHostControl(...)`.

- [ ] **Step 6: Run router tests**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~TrackAnalyticsRoutesPanelEventToAnalyticsService|FullyQualifiedName~TrackAnalyticsRejectsBlankEventName"
```

Expected: PASS.

- [ ] **Step 7: Commit bridge analytics**

```powershell
git add src/OfficeAgent.ExcelAddIn/WebBridge src/OfficeAgent.ExcelAddIn/TaskPane tests/OfficeAgent.ExcelAddIn.Tests/WebMessageRouterTests.cs
git commit -m "feat: route panel analytics through bridge"
```

---

### Task 5: Host Composition And Project File Wiring

**Files:**
- Modify: `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj`
- Test: `tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs`

- [ ] **Step 1: Add source-check test for host analytics composition**

Add to `tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs`:

```csharp
[Fact]
public void ThisAddInComposesAnalyticsServiceFromSettings()
{
    var addInText = File.ReadAllText(ResolveRepositoryPath("src", "OfficeAgent.ExcelAddIn", "ThisAddIn.cs"));

    Assert.Contains("internal IAnalyticsService AnalyticsService { get; private set; }", addInText, StringComparison.Ordinal);
    Assert.Contains("new InsertLogAnalyticsSink(() => SettingsStore.Load())", addInText, StringComparison.Ordinal);
    Assert.Contains("AnalyticsService = string.IsNullOrWhiteSpace(initialSettings.AnalyticsBaseUrl)", addInText, StringComparison.Ordinal);
    Assert.Contains("NoopAnalyticsService.Instance", addInText, StringComparison.Ordinal);
}

[Fact]
public void ExcelAddInProjectIncludesRibbonAnalyticsHelper()
{
    var projectText = File.ReadAllText(ResolveRepositoryPath("src", "OfficeAgent.ExcelAddIn", "OfficeAgent.ExcelAddIn.csproj"));

    Assert.Contains("<Compile Include=\"Analytics\\RibbonAnalyticsHelper.cs\" />", projectText, StringComparison.Ordinal);
}
```

- [ ] **Step 2: Run source-check tests to verify they fail**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~ThisAddInComposesAnalyticsServiceFromSettings|FullyQualifiedName~ExcelAddInProjectIncludesRibbonAnalyticsHelper"
```

Expected: FAIL because host analytics wiring and helper project include do not exist.

- [ ] **Step 3: Compose analytics service in ThisAddIn**

Modify `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs`:

- Add `using OfficeAgent.Core.Analytics;`
- Add `using OfficeAgent.Infrastructure.Analytics;`
- Add property:

```csharp
internal IAnalyticsService AnalyticsService { get; private set; }
```

After `var initialSettings = SettingsStore.Load();`, add:

```csharp
AnalyticsService = string.IsNullOrWhiteSpace(initialSettings.AnalyticsBaseUrl)
    ? NoopAnalyticsService.Instance
    : new AnalyticsService(new InsertLogAnalyticsSink(() => SettingsStore.Load()));
```

Pass `AnalyticsService` into `TaskPaneController` immediately:

```csharp
TaskPaneController = new TaskPaneController(
    this,
    SessionStore,
    SettingsStore,
    ExcelContextService,
    ExcelCommandExecutor,
    AgentOrchestrator,
    SharedCookies,
    CookieStore,
    GetResolvedUiLocale,
    AnalyticsService);
```

Constructor calls for `CurrentBusinessSystemConnector`, `WorksheetSyncService`, `RibbonSyncController`, and `RibbonTemplateController` are updated in the later tasks that add their analytics constructor parameters.

- [ ] **Step 4: Add project include for RibbonAnalyticsHelper**

Modify `src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj` inside the `<ItemGroup>` that contains `<Compile Include="AgentRibbon.cs">`:

```xml
<Compile Include="Analytics\RibbonAnalyticsHelper.cs" />
```

Create the minimal `src/OfficeAgent.ExcelAddIn/Analytics/RibbonAnalyticsHelper.cs` file now so the explicit VSTO project include always points at an existing source file:

```csharp
using OfficeAgent.Core.Analytics;

namespace OfficeAgent.ExcelAddIn.Analytics
{
    internal sealed class RibbonAnalyticsHelper
    {
        private readonly IAnalyticsService analyticsService;

        public RibbonAnalyticsHelper(IAnalyticsService analyticsService)
        {
            this.analyticsService = analyticsService ?? NoopAnalyticsService.Instance;
        }
    }
}
```

- [ ] **Step 5: Run source-check tests**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~ThisAddInComposesAnalyticsServiceFromSettings|FullyQualifiedName~ExcelAddInProjectIncludesRibbonAnalyticsHelper"
```

Expected: PASS.

- [ ] **Step 6: Commit host wiring**

```powershell
git add src/OfficeAgent.ExcelAddIn/ThisAddIn.cs src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs
git commit -m "feat: compose analytics service in add-in host"
```

---

### Task 6: Ribbon Entry And Ribbon Sync Instrumentation

**Files:**
- Create/Modify: `src/OfficeAgent.ExcelAddIn/Analytics/RibbonAnalyticsHelper.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/AgentRibbon.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/RibbonTemplateController.cs`
- Test: `tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs`
- Test: `tests/OfficeAgent.ExcelAddIn.Tests/RibbonSyncControllerTests.cs`

- [ ] **Step 1: Add source-check tests for Ribbon entry events**

Add to `tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs`:

```csharp
[Fact]
public void AgentRibbonTracksPrimaryButtonClicks()
{
    var ribbonText = File.ReadAllText(ResolveRepositoryPath("src", "OfficeAgent.ExcelAddIn", "AgentRibbon.cs"));

    Assert.Contains("TrackRibbonClick(\"ribbon.taskpane.toggle.clicked\"", ribbonText, StringComparison.Ordinal);
    Assert.Contains("TrackRibbonClick(\"ribbon.login.clicked\"", ribbonText, StringComparison.Ordinal);
    Assert.Contains("TrackRibbonClick(\"ribbon.initialize.clicked\"", ribbonText, StringComparison.Ordinal);
    Assert.Contains("TrackRibbonClick(\"ribbon.download.clicked\"", ribbonText, StringComparison.Ordinal);
    Assert.Contains("TrackRibbonClick(\"ribbon.upload.clicked\"", ribbonText, StringComparison.Ordinal);
    Assert.Contains("TrackRibbonClick(\"ribbon.documentation.clicked\"", ribbonText, StringComparison.Ordinal);
    Assert.Contains("TrackRibbonClick(\"ribbon.about.clicked\"", ribbonText, StringComparison.Ordinal);
}
```

- [ ] **Step 2: Add behavioral RibbonSyncController tests**

In `tests/OfficeAgent.ExcelAddIn.Tests/RibbonSyncControllerTests.cs`, update `CreateController` helper to accept `IAnalyticsService analyticsService = null` and pass it to `RibbonSyncController`.

Add:

```csharp
[Fact]
public void ExecuteInitializeCurrentSheetTracksCompletedEventWithProjectName()
{
    var connector = new FakeSystemConnector();
    var analytics = new RecordingAnalyticsService();
    var controller = CreateController(
        connector,
        activeSheetName: "Sheet1",
        analyticsService: analytics);
    controller.SelectProject(new ProjectOption
    {
        SystemKey = connector.SystemKey,
        ProjectId = "performance",
        DisplayName = "绩效项目",
    });

    controller.ExecuteInitializeCurrentSheet();

    Assert.Contains(analytics.Events, analyticsEvent =>
        analyticsEvent.EventName == "ribbon.initialize.completed" &&
        Equals(analyticsEvent.Properties["projectId"], "performance") &&
        Equals(analyticsEvent.Properties["projectName"], "绩效项目") &&
        Equals(analyticsEvent.Properties["sheetName"], "Sheet1"));
}
```

Add this `RecordingAnalyticsService` fake near the other fake classes in that file:

```csharp
private sealed class RecordingAnalyticsService : IAnalyticsService
{
    public List<AnalyticsEvent> Events { get; } = new List<AnalyticsEvent>();

    public void Track(AnalyticsEvent analyticsEvent)
    {
        Events.Add(analyticsEvent);
    }

    public void Track(
        string eventName,
        string source,
        IDictionary<string, object> properties = null,
        IDictionary<string, object> businessContext = null,
        AnalyticsError error = null)
    {
        Events.Add(new AnalyticsEvent
        {
            EventName = eventName,
            Source = source,
            Properties = properties,
            BusinessContext = businessContext,
            Error = error,
        });
    }
}
```

- [ ] **Step 3: Run Ribbon tests to verify they fail**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~AgentRibbonTracksPrimaryButtonClicks|FullyQualifiedName~ExecuteInitializeCurrentSheetTracksCompletedEventWithProjectName"
```

Expected: FAIL because Ribbon analytics calls and constructor injection are missing.

- [ ] **Step 4: Implement RibbonAnalyticsHelper**

Create or replace `src/OfficeAgent.ExcelAddIn/Analytics/RibbonAnalyticsHelper.cs`:

```csharp
using System;
using System.Collections.Generic;
using OfficeAgent.Core.Analytics;
using OfficeAgent.Core.Models;
using OfficeAgent.ExcelAddIn.Localization;

namespace OfficeAgent.ExcelAddIn.Analytics
{
    internal sealed class RibbonAnalyticsHelper
    {
        private readonly IAnalyticsService analyticsService;
        private readonly Func<SheetBinding> getActiveBinding;
        private readonly Func<string> getActiveSheetName;
        private readonly Func<string> getActiveWorkbookName;
        private readonly Func<HostLocalizedStrings> getStrings;

        public RibbonAnalyticsHelper(
            IAnalyticsService analyticsService,
            Func<SheetBinding> getActiveBinding,
            Func<string> getActiveSheetName,
            Func<string> getActiveWorkbookName,
            Func<HostLocalizedStrings> getStrings)
        {
            this.analyticsService = analyticsService ?? NoopAnalyticsService.Instance;
            this.getActiveBinding = getActiveBinding ?? (() => null);
            this.getActiveSheetName = getActiveSheetName ?? (() => string.Empty);
            this.getActiveWorkbookName = getActiveWorkbookName ?? (() => string.Empty);
            this.getStrings = getStrings ?? (() => HostLocalizedStrings.ForLocale("en"));
        }

        public void Track(string eventName, IDictionary<string, object> properties = null, AnalyticsError error = null)
        {
            var merged = BuildCommonProperties();
            if (properties != null)
            {
                foreach (var item in properties)
                {
                    merged[item.Key] = item.Value;
                }
            }

            analyticsService.Track(eventName, "ribbon", merged, error: error);
        }

        private Dictionary<string, object> BuildCommonProperties()
        {
            var binding = getActiveBinding();
            var strings = getStrings();
            return new Dictionary<string, object>(StringComparer.Ordinal)
            {
                ["systemKey"] = binding?.SystemKey ?? string.Empty,
                ["projectId"] = binding?.ProjectId ?? string.Empty,
                ["projectName"] = binding?.ProjectName ?? string.Empty,
                ["sheetName"] = binding?.SheetName ?? getActiveSheetName(),
                ["workbookName"] = getActiveWorkbookName(),
                ["uiLocale"] = strings?.Locale ?? string.Empty,
            };
        }
    }
}
```

`HostLocalizedStrings` already exposes `Locale`, so the helper can use it directly for `uiLocale`.

- [ ] **Step 5: Instrument AgentRibbon click handlers**

Modify `src/OfficeAgent.ExcelAddIn/AgentRibbon.cs`:

- Add `using OfficeAgent.ExcelAddIn.Analytics;`
- Add field:

```csharp
private RibbonAnalyticsHelper analytics;
```

- Initialize it in `AgentRibbon_Load` or `BindToControllersAndRefresh`:

```csharp
analytics = Globals.ThisAddIn?.CreateRibbonAnalyticsHelper();
```

Add this factory method to `ThisAddIn`:

```csharp
internal RibbonAnalyticsHelper CreateRibbonAnalyticsHelper()
{
    return new RibbonAnalyticsHelper(
        AnalyticsService,
        () => RibbonSyncController?.ActiveBinding,
        GetActiveWorksheetName,
        GetActiveWorkbookName,
        () => HostLocalizedStrings);
}
```

Add `internal SheetBinding ActiveBinding { get; private set; }` to `RibbonSyncController`.

In `ApplyBindingState(SheetBinding binding)`, set:

```csharp
ActiveBinding = binding;
```

before updating `ActiveProjectId`.

In `ClearActiveProjectState()`, set:

```csharp
ActiveBinding = null;
```

before clearing `ActiveProjectId`.

Add `private string GetActiveWorkbookName()` to `ThisAddIn`:

```csharp
private string GetActiveWorkbookName()
{
    try
    {
        var workbook = Application?.ActiveWorkbook;
        return workbook?.Name ?? string.Empty;
    }
    catch
    {
        return string.Empty;
    }
}
```

- Add helper:

```csharp
private void TrackRibbonClick(string eventName, IDictionary<string, object> properties = null)
{
    analytics?.Track(eventName, properties);
}
```

- In handlers add first-line calls:

```csharp
TrackRibbonClick("ribbon.taskpane.toggle.clicked");
TrackRibbonClick("ribbon.login.clicked");
TrackRibbonClick("ribbon.initialize.clicked");
TrackRibbonClick("ribbon.ai_map_columns.clicked");
TrackRibbonClick("ribbon.download.clicked", new Dictionary<string, object> { ["operation"] = "partialDownload" });
TrackRibbonClick("ribbon.upload.clicked", new Dictionary<string, object> { ["operation"] = "partialUpload" });
TrackRibbonClick("ribbon.documentation.clicked", new Dictionary<string, object> { ["url"] = DocumentationUrl });
TrackRibbonClick("ribbon.about.clicked", new Dictionary<string, object> { ["version"] = VersionInfo.AppVersion });
```

Also track project dropdown:

```csharp
TrackRibbonClick("ribbon.project_dropdown.opened");
```

inside `ProjectDropDown_ItemsLoading`, and:

```csharp
TrackRibbonClick("ribbon.project.selected", new Dictionary<string, object>
{
    ["projectSelectionKey"] = selectedKey,
    ["projectName"] = selectedProject.DisplayName ?? string.Empty,
});
```

inside successful `ProjectDropDown_SelectionChanged`.

- [ ] **Step 6: Instrument RibbonSyncController result paths**

Modify `src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs`:

- Add `IAnalyticsService analyticsService = null` constructor parameter and field.
- Use `analyticsService ?? NoopAnalyticsService.Instance`.
- Add private `TrackRibbonEvent(...)` that merges current binding/project fields:

```csharp
private void TrackRibbonEvent(string eventName, IDictionary<string, object> properties = null, AnalyticsError error = null)
{
    var merged = new Dictionary<string, object>(StringComparer.Ordinal)
    {
        ["systemKey"] = ActiveBinding?.SystemKey ?? string.Empty,
        ["projectId"] = ActiveBinding?.ProjectId ?? ActiveProjectId ?? string.Empty,
        ["projectName"] = ActiveBinding?.ProjectName ?? ActiveProjectDisplayName ?? string.Empty,
        ["sheetName"] = ActiveBinding?.SheetName ?? GetRequiredSheetName(),
        ["uiLocale"] = GetStrings().Locale,
    };

    if (properties != null)
    {
        foreach (var item in properties)
        {
            merged[item.Key] = item.Value;
        }
    }

    analyticsService.Track(eventName, "ribbon", merged, error: error);
}
```
- In `SelectProject`, record:
  - `ribbon.project_layout.confirmed`
  - `ribbon.project_layout.canceled`
  - `ribbon.project.selected`
- In `ExecuteInitializeCurrentSheet`, record:
  - `ribbon.initialize.completed`
  - `ribbon.initialize.failed`
- In `ExecuteAiColumnMapping`, record:
  - `ribbon.ai_map_columns.completed`
  - `ribbon.ai_map_columns.failed`
- In `ExecuteDownload`, record:
  - `ribbon.download.confirmed`
  - `ribbon.download.canceled`
  - `ribbon.download.completed`
  - `ribbon.download.failed`
- In `ExecuteUpload`, record:
  - `ribbon.upload.previewed`
  - `ribbon.upload.confirmed`
  - `ribbon.upload.canceled`
  - `ribbon.upload.completed`
  - `ribbon.upload.failed`

For failures use:

```csharp
error: new AnalyticsError
{
    Code = "operation_failed",
    Message = ex.Message,
    ExceptionType = ex.GetType().Name,
}
```

- [ ] **Step 7: Instrument RibbonTemplateController**

Modify `src/OfficeAgent.ExcelAddIn/RibbonTemplateController.cs`:

- Add `using System.Collections.Generic;` and `using OfficeAgent.Core.Analytics;`.
- Add optional `IAnalyticsService analyticsService = null` to the internal constructor and assign `this.analyticsService = analyticsService ?? NoopAnalyticsService.Instance;`.
- Update the public constructor to call the internal constructor with `NoopAnalyticsService.Instance`.
- Track:
  - `ribbon.template.apply.clicked`
  - `ribbon.template.apply.completed`
  - `ribbon.template.apply.failed`
  - `ribbon.template.save.clicked`
  - `ribbon.template.save.completed`
  - `ribbon.template.save.failed`
  - `ribbon.template.save_as.clicked`
  - `ribbon.template.save_as.completed`
  - `ribbon.template.save_as.failed`

Add helper:

```csharp
private void TrackTemplateEvent(string eventName, string sheetName, SheetTemplateState state, IDictionary<string, object> properties = null, AnalyticsError error = null)
{
    var merged = new Dictionary<string, object>(StringComparer.Ordinal)
    {
        ["sheetName"] = sheetName ?? string.Empty,
        ["projectName"] = state?.ProjectDisplayName ?? string.Empty,
        ["templateId"] = state?.TemplateId ?? string.Empty,
        ["templateName"] = state?.TemplateName ?? string.Empty,
        ["templateRevision"] = state?.TemplateRevision ?? 0,
    };

    if (properties != null)
    {
        foreach (var item in properties)
        {
            merged[item.Key] = item.Value;
        }
    }

    analyticsService.Track(eventName, "ribbon", merged, error: error);
}
```

Call it at the start of each command with `*.clicked`, after successful catalog mutation with `*.completed`, and in each catch block with `*.failed` plus:

```csharp
new AnalyticsError
{
    Code = "template_operation_failed",
    Message = ex.Message,
    ExceptionType = ex.GetType().Name,
}
```

- [ ] **Step 8: Run Ribbon tests**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~AgentRibbonTracksPrimaryButtonClicks|FullyQualifiedName~ExecuteInitializeCurrentSheetTracksCompletedEventWithProjectName"
```

Expected: PASS.

- [ ] **Step 9: Commit Ribbon instrumentation**

```powershell
git add src/OfficeAgent.ExcelAddIn/Analytics/RibbonAnalyticsHelper.cs src/OfficeAgent.ExcelAddIn/AgentRibbon.cs src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs src/OfficeAgent.ExcelAddIn/RibbonTemplateController.cs src/OfficeAgent.ExcelAddIn/ThisAddIn.cs tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs tests/OfficeAgent.ExcelAddIn.Tests/RibbonSyncControllerTests.cs
git commit -m "feat: instrument ribbon analytics"
```

---

### Task 7: Connector Analytics And Business Context

**Files:**
- Modify: `src/OfficeAgent.Core/Sync/WorksheetSyncService.cs`
- Modify: `src/OfficeAgent.Infrastructure/Http/CurrentBusinessSystemConnector.cs`
- Test: `tests/OfficeAgent.Core.Tests/WorksheetSyncServiceTests.cs`
- Test: `tests/OfficeAgent.Infrastructure.Tests/CurrentBusinessSystemConnectorTests.cs`

- [ ] **Step 1: Add WorksheetSyncService tests**

In `tests/OfficeAgent.Core.Tests/WorksheetSyncServiceTests.cs`, add `using OfficeAgent.Core.Analytics;`.

Add:

```csharp
[Fact]
public void DownloadTracksConnectorCompletedEvent()
{
    var connector = new FakeSystemConnector();
    var analytics = new RecordingAnalyticsService();
    var service = CreateService(connector, analyticsService: analytics);

    service.Download(connector.SystemKey, "performance", new[] { "row-1" }, new[] { "owner_name" });

    Assert.Contains(analytics.Events, analyticsEvent =>
        analyticsEvent.EventName == "connector.find.completed" &&
        Equals(analyticsEvent.Properties["systemKey"], connector.SystemKey) &&
        Equals(analyticsEvent.Properties["projectId"], "performance") &&
        Equals(analyticsEvent.Properties["rowIdCount"], 1) &&
        Equals(analyticsEvent.Properties["fieldKeyCount"], 1));
}

[Fact]
public void UploadTracksConnectorFailedEventWhenBatchSaveThrows()
{
    var connector = new FakeSystemConnector
    {
        BatchSaveException = new InvalidOperationException("save failed"),
    };
    var analytics = new RecordingAnalyticsService();
    var service = CreateService(connector, analyticsService: analytics);

    Assert.Throws<InvalidOperationException>(() =>
        service.Upload(connector.SystemKey, "performance", new[] { new CellChange { RowId = "row-1", ApiFieldKey = "owner_name", NewValue = "李四" } }));

    Assert.Contains(analytics.Events, analyticsEvent =>
        analyticsEvent.EventName == "connector.batch_save.failed" &&
        analyticsEvent.Error != null &&
        analyticsEvent.Error.Message == "save failed");
}
```

Update `CreateService` helper to accept `IAnalyticsService analyticsService = null`.

- [ ] **Step 2: Add CurrentBusinessSystemConnector businessContext test**

In `tests/OfficeAgent.Infrastructure.Tests/CurrentBusinessSystemConnectorTests.cs`, add:

```csharp
[Fact]
public void FindTracksBusinessContextWithoutRawPayload()
{
    var handler = new RecordingHandler(request => new HttpResponseMessage(HttpStatusCode.OK)
    {
        Content = new StringContent("[{\"row_id\":\"row-1\",\"owner_name\":\"张三\"}]"),
    });
    var analytics = new RecordingAnalyticsService();
    var connector = CurrentBusinessSystemConnector.ForTests(
        "https://api.internal.example",
        handler,
        analytics);

    connector.Find("performance", new[] { "row-1" }, new[] { "owner_name" });

    Assert.Contains(analytics.Events, analyticsEvent =>
        analyticsEvent.EventName == "business.current.find.completed" &&
        Equals(analyticsEvent.BusinessContext["endpoint"], "/find") &&
        Equals(analyticsEvent.Properties["projectId"], "performance"));
    Assert.DoesNotContain(analytics.Events, analyticsEvent =>
        analyticsEvent.BusinessContext.ContainsKey("requestBody") ||
        analyticsEvent.BusinessContext.ContainsKey("responseBody"));
}
```

Overload `CurrentBusinessSystemConnector.ForTests` in the implementation to support the third `IAnalyticsService` argument.

- [ ] **Step 3: Run connector tests to verify they fail**

Run:

```powershell
dotnet test tests/OfficeAgent.Core.Tests/OfficeAgent.Core.Tests.csproj --filter "FullyQualifiedName~DownloadTracksConnectorCompletedEvent|FullyQualifiedName~UploadTracksConnectorFailedEventWhenBatchSaveThrows"
dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj --filter FullyQualifiedName~FindTracksBusinessContextWithoutRawPayload
```

Expected: FAIL because analytics injection and events are missing.

- [ ] **Step 4: Instrument WorksheetSyncService**

Modify `src/OfficeAgent.Core/Sync/WorksheetSyncService.cs`:

- Add `using System.Diagnostics;`
- Add `using OfficeAgent.Core.Analytics;`
- Add field:

```csharp
private readonly IAnalyticsService analyticsService;
```

- Update constructors with optional analytics:

```csharp
public WorksheetSyncService(
    ISystemConnectorRegistry connectorRegistry,
    IWorksheetMetadataStore metadataStore,
    IAnalyticsService analyticsService = null)
```

and assign `this.analyticsService = analyticsService ?? NoopAnalyticsService.Instance;`.

- Wrap `Download`:

```csharp
public IReadOnlyList<IDictionary<string, object>> Download(
    string systemKey,
    string projectId,
    IReadOnlyList<string> rowIds,
    IReadOnlyList<string> fieldKeys)
{
    var stopwatch = Stopwatch.StartNew();
    var properties = BuildConnectorProperties(systemKey, projectId);
    properties["rowIdCount"] = rowIds?.Count ?? 0;
    properties["fieldKeyCount"] = fieldKeys?.Count ?? 0;
    try
    {
        var rows = GetRequiredConnector(systemKey).Find(projectId, rowIds, fieldKeys);
        stopwatch.Stop();
        properties["durationMs"] = stopwatch.ElapsedMilliseconds;
        properties["resultCount"] = rows?.Count ?? 0;
        analyticsService.Track("connector.find.completed", "connector", properties);
        return rows;
    }
    catch (Exception ex)
    {
        stopwatch.Stop();
        properties["durationMs"] = stopwatch.ElapsedMilliseconds;
        analyticsService.Track("connector.find.failed", "connector", properties, error: ToAnalyticsError(ex));
        throw;
    }
}
```

Wrap the remaining methods with these event names and properties:

| Method | Success Event | Failure Event | Extra Properties |
| --- | --- | --- | --- |
| `GetProjects` | `connector.projects.completed` | `connector.projects.failed` | `projectCount` |
| `InitializeSheet` | `connector.initialize_sheet.completed` | `connector.initialize_sheet.failed` | `sheetName`, `projectName` |
| `CreateBindingSeed` | `connector.binding_seed.completed` | `connector.binding_seed.failed` | `sheetName`, `projectName` |
| `LoadFieldMappingDefinition` | `connector.field_mapping_definition.completed` | `connector.field_mapping_definition.failed` | none |
| `Upload` | `connector.batch_save.completed` | `connector.batch_save.failed` | `changeCount` |
| `FilterUploadChanges` | `connector.upload_filter.completed` | `connector.upload_filter.failed` | `changeCount`, `includedCount`, `skippedCount` |

Each wrapper should start a `Stopwatch`, add `durationMs` before success or failure tracking, and rethrow the original exception after tracking failure.

- Add helpers:

```csharp
private static Dictionary<string, object> BuildConnectorProperties(string systemKey, string projectId)
{
    return new Dictionary<string, object>(StringComparer.Ordinal)
    {
        ["systemKey"] = systemKey ?? string.Empty,
        ["projectId"] = projectId ?? string.Empty,
    };
}

private static AnalyticsError ToAnalyticsError(Exception ex)
{
    return new AnalyticsError
    {
        Code = "connector_failed",
        Message = ex.Message,
        ExceptionType = ex.GetType().Name,
    };
}
```

- [ ] **Step 5: Instrument CurrentBusinessSystemConnector**

Modify `src/OfficeAgent.Infrastructure/Http/CurrentBusinessSystemConnector.cs`:

- Add `using System.Diagnostics;`
- Add `using OfficeAgent.Core.Analytics;`
- Add field:

```csharp
private readonly IAnalyticsService analyticsService;
```

- Add optional constructor parameter `IAnalyticsService analyticsService = null` to public and private constructors and assign `NoopAnalyticsService.Instance`.
- Add `ForTests(string baseUrl, HttpMessageHandler handler, IAnalyticsService analyticsService)` overload.
- Around `Find`, track:

```csharp
TrackBusinessEvent("business.current.find.completed", projectId, "/find", stopwatch.ElapsedMilliseconds, new Dictionary<string, object>
{
    ["rowIdCount"] = requestedRowIds.Count,
    ["fieldKeyCount"] = fieldKeys?.Count ?? 0,
});
```

- Around `BatchSave`, track `business.current.batch_save.completed` with `changeCount`.
- Around `GetProjects`, track `business.current.projects.completed` / `business.current.projects.failed` with endpoint `/projects` and `projectCount`.
- Around `GetSchema`, track `business.current.schema.completed` / `business.current.schema.failed` with endpoint `/head+/find`.
- Around `BuildFieldMappingSeed`, track `business.current.field_mapping_seed.completed` / `business.current.field_mapping_seed.failed` with endpoint `/head+/find` and `fieldMappingCount`.

Helper:

```csharp
private void TrackBusinessEvent(
    string eventName,
    string projectId,
    string endpoint,
    long durationMs,
    IDictionary<string, object> properties = null,
    AnalyticsError error = null)
{
    var merged = new Dictionary<string, object>(StringComparer.Ordinal)
    {
        ["systemKey"] = CurrentSystemKey,
        ["projectId"] = projectId ?? string.Empty,
        ["durationMs"] = durationMs,
    };
    if (properties != null)
    {
        foreach (var item in properties)
        {
            merged[item.Key] = item.Value;
        }
    }

    analyticsService.Track(
        eventName,
        "connector",
        merged,
        new Dictionary<string, object>(StringComparer.Ordinal)
        {
            ["endpoint"] = endpoint ?? string.Empty,
            ["module"] = "current-business-system",
        },
        error);
}
```

- [ ] **Step 6: Run connector tests**

Run:

```powershell
dotnet test tests/OfficeAgent.Core.Tests/OfficeAgent.Core.Tests.csproj --filter "FullyQualifiedName~DownloadTracksConnectorCompletedEvent|FullyQualifiedName~UploadTracksConnectorFailedEventWhenBatchSaveThrows"
dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj --filter FullyQualifiedName~FindTracksBusinessContextWithoutRawPayload
```

Expected: PASS.

- [ ] **Step 7: Commit connector analytics**

```powershell
git add src/OfficeAgent.Core/Sync/WorksheetSyncService.cs src/OfficeAgent.Infrastructure/Http/CurrentBusinessSystemConnector.cs tests/OfficeAgent.Core.Tests/WorksheetSyncServiceTests.cs tests/OfficeAgent.Infrastructure.Tests/CurrentBusinessSystemConnectorTests.cs
git commit -m "feat: instrument connector analytics"
```

---

### Task 8: Frontend Bridge Types And Native Bridge

**Files:**
- Modify: `src/OfficeAgent.Frontend/src/types/bridge.ts`
- Modify: `src/OfficeAgent.Frontend/src/bridge/nativeBridge.ts`
- Test: `src/OfficeAgent.Frontend/src/bridge/nativeBridge.test.ts`

- [ ] **Step 1: Add failing nativeBridge tests**

Append to `src/OfficeAgent.Frontend/src/bridge/nativeBridge.test.ts`:

```typescript
it('posts trackAnalytics requests to the native bridge', async () => {
  const webView = createMockWebView();
  const bridge = new NativeBridge(webView);
  const promise = bridge.trackAnalytics({
    eventName: 'panel.composer.send.clicked',
    source: 'panel',
    properties: { inputLength: 12 },
    businessContext: { module: 'demo' },
  });

  const request = webView.postedMessages[0];
  expect(request.type).toBe('bridge.trackAnalytics');
  expect(request.payload.eventName).toBe('panel.composer.send.clicked');

  webView.dispatch({
    type: request.type,
    requestId: request.requestId,
    ok: true,
    payload: { tracked: true },
  });

  await expect(promise).resolves.toEqual({ tracked: true });
});

it('resolves trackAnalytics in browser preview without a native bridge', async () => {
  const bridge = new NativeBridge(undefined);

  await expect(bridge.trackAnalytics({
    eventName: 'panel.opened',
    source: 'panel',
  })).resolves.toEqual({ tracked: false });
});
```

Use the existing `createMockWebView()` helper in that file.

- [ ] **Step 2: Run nativeBridge tests to verify they fail**

Run:

```powershell
cd src/OfficeAgent.Frontend
npm run test -- src/bridge/nativeBridge.test.ts
```

Expected: FAIL because `trackAnalytics` and analytics payload types do not exist.

- [ ] **Step 3: Add frontend analytics types**

Modify `src/OfficeAgent.Frontend/src/types/bridge.ts`:

```typescript
export interface AnalyticsPayload {
  eventName: string;
  source?: 'panel' | 'bridge' | 'ribbon' | 'connector' | 'business' | 'host';
  properties?: Record<string, unknown>;
  businessContext?: Record<string, unknown>;
}

export interface AnalyticsResult {
  tracked: boolean;
}
```

Add to `AppSettings`:

```typescript
analyticsBaseUrl: string;
```

- [ ] **Step 4: Add nativeBridge.trackAnalytics**

Modify `src/OfficeAgent.Frontend/src/bridge/nativeBridge.ts`:

- Import `AnalyticsPayload` and `AnalyticsResult`.
- Add bridge type:

```typescript
trackAnalytics: 'bridge.trackAnalytics',
```

- Add `analyticsBaseUrl: ''` to `BROWSER_PREVIEW_SETTINGS`.
- Preserve `analyticsBaseUrl` in browser-preview `saveSettings`.
- Add method:

```typescript
trackAnalytics(payload: AnalyticsPayload) {
  return this.invoke<AnalyticsPayload, AnalyticsResult>(BRIDGE_TYPES.trackAnalytics, payload);
}
```

- In browser-preview branch, add:

```typescript
if (type === BRIDGE_TYPES.trackAnalytics) {
  return Promise.resolve({ tracked: false } as TResult);
}
```

- [ ] **Step 5: Run nativeBridge tests**

Run:

```powershell
cd src/OfficeAgent.Frontend
npm run test -- src/bridge/nativeBridge.test.ts
```

Expected: PASS.

- [ ] **Step 6: Commit frontend bridge analytics**

```powershell
git add src/OfficeAgent.Frontend/src/types/bridge.ts src/OfficeAgent.Frontend/src/bridge/nativeBridge.ts src/OfficeAgent.Frontend/src/bridge/nativeBridge.test.ts
git commit -m "feat: add panel analytics bridge client"
```

---

### Task 9: Panel Settings Field And Interaction Events

**Files:**
- Create: `src/OfficeAgent.Frontend/src/analytics/panelAnalytics.ts`
- Modify: `src/OfficeAgent.Frontend/src/App.tsx`
- Modify: `src/OfficeAgent.Frontend/src/i18n/uiStrings.ts`
- Test: `src/OfficeAgent.Frontend/src/App.test.tsx`

- [ ] **Step 1: Add failing App tests**

In `src/OfficeAgent.Frontend/src/App.test.tsx`, first add `trackAnalytics: vi.fn(),` to the mocked `nativeBridge` object and add this default in `beforeEach`:

```typescript
mockedBridge.trackAnalytics.mockResolvedValue({ tracked: true });
```

Then add these tests:

```typescript
it('saves the analytics base URL setting', async () => {
  const user = userEvent.setup();
  render(<App />);

  await user.click(screen.getByLabelText(/打开设置|open settings/i));
  await user.clear(screen.getByLabelText(/埋点 Base URL|Analytics Base URL/i));
  await user.type(screen.getByLabelText(/埋点 Base URL|Analytics Base URL/i), 'http://localhost:3200');
  await user.click(screen.getByRole('button', { name: /保存|save/i }));

  expect(mockedBridge.saveSettings).toHaveBeenCalledWith(expect.objectContaining({
    analyticsBaseUrl: 'http://localhost:3200',
  }));
});

it('tracks composer send events without sending prompt text', async () => {
  const user = userEvent.setup();
  render(<App />);

  await user.type(screen.getByRole('textbox', { name: /消息输入框|message composer/i }), 'Create a summary sheet');
  await user.click(screen.getByRole('button', { name: /发送|send/i }));

  expect(mockedBridge.trackAnalytics).toHaveBeenCalledWith(expect.objectContaining({
    eventName: 'panel.composer.send.clicked',
    properties: expect.objectContaining({
      inputLength: 22,
    }),
  }));
  expect(JSON.stringify(mockedBridge.trackAnalytics.mock.calls)).not.toContain('Create a summary sheet');
});
```

The key assertions are `analyticsBaseUrl`, `panel.composer.send.clicked`, and no raw prompt text.

- [ ] **Step 2: Run App tests to verify they fail**

Run:

```powershell
cd src/OfficeAgent.Frontend
npm run test -- src/App.test.tsx
```

Expected: FAIL because settings field and panel analytics helper are missing.

- [ ] **Step 3: Add panel analytics helper**

Create `src/OfficeAgent.Frontend/src/analytics/panelAnalytics.ts`:

```typescript
import { nativeBridge } from '../bridge/nativeBridge';
import type { AnalyticsPayload, SelectionContext, UiLocale } from '../types/bridge';

type CommonContext = {
  sessionId?: string;
  uiLocale: UiLocale;
  selectionContext?: SelectionContext | null;
};

export function trackPanelEvent(
  eventName: string,
  common: CommonContext,
  properties: Record<string, unknown> = {},
  businessContext: Record<string, unknown> = {},
) {
  const payload: AnalyticsPayload = {
    eventName,
    source: 'panel',
    properties: {
      sessionId: common.sessionId ?? '',
      uiLocale: common.uiLocale,
      hasSelection: Boolean(common.selectionContext?.hasSelection),
      sheetName: common.selectionContext?.sheetName ?? '',
      workbookName: common.selectionContext?.workbookName ?? '',
      ...properties,
    },
    businessContext,
  };

  void nativeBridge.trackAnalytics(payload).catch(() => {
    // Analytics must never break the panel interaction.
  });
}
```

- [ ] **Step 4: Add Analytics Base URL UI strings**

Modify `src/OfficeAgent.Frontend/src/i18n/uiStrings.ts`.

Add keys:

```typescript
analyticsBaseUrlFieldLabel: 'Analytics Base URL',
```

and Chinese:

```typescript
analyticsBaseUrlFieldLabel: '埋点 Base URL',
```

- [ ] **Step 5: Update App settings and interaction handlers**

Modify `src/OfficeAgent.Frontend/src/App.tsx`:

- Add `analyticsBaseUrl: ''` to `DEFAULT_SETTINGS`.
- Add field after Business Base URL:

```tsx
<label className="settings-field">
  <span>{strings.analyticsBaseUrlFieldLabel}</span>
  <input
    aria-label={strings.analyticsBaseUrlFieldLabel}
    type="url"
    value={draftSettings.analyticsBaseUrl}
    onChange={(event) => updateDraftSettings({ analyticsBaseUrl: event.target.value })}
  />
</label>
```

- Import `trackPanelEvent`.
- Add helper:

```typescript
function track(eventName: string, properties: Record<string, unknown> = {}, businessContext: Record<string, unknown> = {}) {
  trackPanelEvent(eventName, {
    sessionId: activeSession?.id,
    uiLocale,
    selectionContext,
  }, properties, businessContext);
}
```

- Add calls:

```typescript
track('panel.settings.opened');
track('panel.settings.saved', {
  apiFormat: savedSettings.apiFormat,
  hasBaseUrl: Boolean(savedSettings.baseUrl.trim()),
  hasBusinessBaseUrl: Boolean(savedSettings.businessBaseUrl.trim()),
  hasAnalyticsBaseUrl: Boolean(savedSettings.analyticsBaseUrl.trim()),
});
track('panel.settings.save_failed', {}, {},);
track('panel.composer.send.clicked', {
  inputLength: trimmedValue.length,
  commandType: command?.commandType ?? (matchesDirectSkillInput(trimmedValue) ? 'skill.upload_data' : 'agent'),
});
track('panel.confirmation.confirmed', { confirmationKind: activePendingConfirmation.kind, commandType: activePendingConfirmation.command?.commandType ?? '' });
track('panel.confirmation.canceled', { confirmationKind: activePendingConfirmation.kind, commandType: activePendingConfirmation.command?.commandType ?? '' });
```

Do not include `trimmedValue` or any raw prompt text in analytics properties.

- [ ] **Step 6: Run App tests**

Run:

```powershell
cd src/OfficeAgent.Frontend
npm run test -- src/App.test.tsx
```

Expected: PASS.

- [ ] **Step 7: Commit panel analytics**

```powershell
git add src/OfficeAgent.Frontend/src/analytics/panelAnalytics.ts src/OfficeAgent.Frontend/src/App.tsx src/OfficeAgent.Frontend/src/i18n/uiStrings.ts src/OfficeAgent.Frontend/src/App.test.tsx
git commit -m "feat: instrument panel analytics"
```

---

### Task 10: Mock Server Analytics Endpoints

**Files:**
- Modify: `tests/mock-server/server.js`
- Modify: `tests/mock-server/README.md`

- [ ] **Step 1: Add mock server analytics routes**

Modify `tests/mock-server/server.js`.

Near `const uploadedProjects = {};`, add:

```javascript
const analyticsLogs = [];
const maxAnalyticsLogs = 500;
```

After `apiApp.use(cookieParser());`, add:

```javascript
apiApp.post("/insertLog", function (req, res) {
  var payload = req.body || {};
  if (payload.frontEndIntent !== "excelAi" || payload.clientSource !== "Excel" || payload.questionType !== 1) {
    return res.status(400).json({ code: "bad_request", message: "insertLog 固定字段不正确。" });
  }

  if (typeof payload.askId !== "string" || !payload.askId.trim() ||
      typeof payload.talkId !== "string" || !payload.talkId.trim() ||
      typeof payload.answer !== "string" || !payload.answer.trim()) {
    return res.status(400).json({ code: "bad_request", message: "askId、talkId、answer 字段必填。" });
  }

  var parsedAnswer;
  try {
    parsedAnswer = JSON.parse(payload.answer);
  } catch (_error) {
    return res.status(400).json({ code: "bad_request", message: "answer 必须是 JSON 字符串。" });
  }

  analyticsLogs.push({
    receivedAtUtc: new Date().toISOString(),
    payload: payload,
    answer: parsedAnswer,
  });

  while (analyticsLogs.length > maxAnalyticsLogs) {
    analyticsLogs.shift();
  }

  res.json({ ok: true, count: analyticsLogs.length });
});

apiApp.get("/analytics/logs", function (_req, res) {
  res.json({ count: analyticsLogs.length, logs: analyticsLogs });
});

apiApp.delete("/analytics/logs", function (_req, res) {
  analyticsLogs.length = 0;
  res.json({ ok: true, count: 0 });
});
```

Keep `/insertLog` outside `requireAuth`.

- [ ] **Step 2: Update mock server README**

Modify `tests/mock-server/README.md`:

- Add config bullet:

```markdown
- `Analytics Base URL = http://localhost:3200`
```

- Add section:

```markdown
### Analytics / 埋点接口

#### `POST /insertLog`

模拟内网埋点接口。请求体固定字段为 `frontEndIntent = excelAi`、`clientSource = Excel`、`questionType = 1`，`answer` 必须是可解析 JSON 字符串。

#### `GET /analytics/logs`

返回当前进程内最近 500 条埋点记录，包含原始外层 payload 和解析后的 `answer`。

#### `DELETE /analytics/logs`

清空当前进程内埋点记录。
```

- [ ] **Step 3: Manually verify mock endpoint**

Run:

```powershell
cd tests/mock-server
npm start
```

In a second terminal:

```powershell
Invoke-RestMethod -Method Post -Uri http://localhost:3200/insertLog -ContentType 'application/json' -Body '{"frontEndIntent":"excelAi","clientSource":"Excel","questionType":1,"askId":"ask-test","talkId":"talk-test","answer":"{\"eventName\":\"panel.opened\",\"schemaVersion\":1}"}'
Invoke-RestMethod -Method Get -Uri http://localhost:3200/analytics/logs
```

Expected: first command returns `{ ok: true, count: 1 }`; second includes `answer.eventName = panel.opened`.

- [ ] **Step 4: Commit mock server changes**

```powershell
git add tests/mock-server/server.js tests/mock-server/README.md
git commit -m "feat: add analytics mock endpoint"
```

---

### Task 11: Documentation Updates

**Files:**
- Modify: `docs/modules/task-pane-current-behavior.md`
- Modify: `docs/modules/ribbon-sync-current-behavior.md`
- Modify: `docs/module-index.md`
- Modify: `docs/vsto-manual-test-checklist.md`

- [ ] **Step 1: Update Task Pane behavior doc**

Modify `docs/modules/task-pane-current-behavior.md`:

- In settings list, add:

```markdown
- `Analytics Base URL`：运营埋点 `/insertLog` 接口 endpoint；为空时不发送埋点
```

- Add a short section:

```markdown
## Panel Analytics

任务窗格会通过 `bridge.trackAnalytics` 上报关键用户交互，例如发送、确认/取消确认卡片、设置保存和会话操作。事件只包含结构化维度，例如输入长度、命令类型、会话 ID、选区摘要和 UI 语言；不会上报用户输入全文、API key、cookie 或单元格原始值。
```

- [ ] **Step 2: Update Ribbon Sync behavior doc**

Modify `docs/modules/ribbon-sync-current-behavior.md`:

Add near Ribbon entry section:

```markdown
Ribbon Sync 会记录运营埋点事件，用于统计项目选择、初始化、AI 映射列、下载、上传、配置按钮和结果状态。事件包含 `systemKey`、`projectId`、`projectName`、`sheetName`、操作类型、确认/取消/成功/失败状态等低敏维度；不会上报单元格原始业务值。
```

- [ ] **Step 3: Update module index**

Modify `docs/module-index.md` to add Analytics as a related design under Task Pane and Ribbon Sync, or add a new row:

```markdown
| Analytics | [docs/superpowers/specs/2026-05-12-office-agent-analytics-instrumentation-design.md](./superpowers/specs/2026-05-12-office-agent-analytics-instrumentation-design.md) | [docs/superpowers/plans/2026-05-12-office-agent-analytics-instrumentation-implementation-plan.md](./superpowers/plans/2026-05-12-office-agent-analytics-instrumentation-implementation-plan.md) | [tests/mock-server/README.md](../tests/mock-server/README.md) |
```

- [ ] **Step 4: Update manual checklist**

Modify `docs/vsto-manual-test-checklist.md`:

Add:

```markdown
## Analytics

- Start `tests/mock-server` and set `Analytics Base URL = http://localhost:3200`.
- Clear existing events with `DELETE http://localhost:3200/analytics/logs`.
- Click Ribbon initialize, download, and upload; confirm `/analytics/logs` contains `ribbon.initialize.*`, `ribbon.download.*`, and `ribbon.upload.*` events with `projectId` and `projectName`.
- In the task pane, send a prompt, save settings, and confirm/cancel a preview card; confirm `/analytics/logs` contains `panel.*` events.
- Confirm logged analytics entries do not contain API keys, cookies, raw prompt text, or cell values.
```

- [ ] **Step 5: Commit docs**

```powershell
git add docs/modules/task-pane-current-behavior.md docs/modules/ribbon-sync-current-behavior.md docs/module-index.md docs/vsto-manual-test-checklist.md
git commit -m "docs: document analytics instrumentation"
```

---

### Task 12: Full Verification

**Files:**
- No source edits expected.

- [ ] **Step 1: Run Core tests**

Run:

```powershell
dotnet test tests/OfficeAgent.Core.Tests/OfficeAgent.Core.Tests.csproj
```

Expected: PASS.

- [ ] **Step 2: Run Infrastructure tests**

Run:

```powershell
dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj
```

Expected: PASS.

- [ ] **Step 3: Run ExcelAddIn tests**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj
```

Expected: PASS. Run the build first so reflection-based tests load the current add-in assembly:

```powershell
dotnet build src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj
```

- [ ] **Step 4: Run frontend build and tests**

Run:

```powershell
cd src/OfficeAgent.Frontend
npm run test
npm run build
```

Expected: PASS.

- [ ] **Step 5: Run mock server smoke test**

Run:

```powershell
cd tests/mock-server
npm start
```

In another terminal:

```powershell
Invoke-RestMethod -Method Delete -Uri http://localhost:3200/analytics/logs
Invoke-RestMethod -Method Post -Uri http://localhost:3200/insertLog -ContentType 'application/json' -Body '{"frontEndIntent":"excelAi","clientSource":"Excel","questionType":1,"askId":"ask-test","talkId":"talk-test","answer":"{\"schemaVersion\":1,\"eventName\":\"panel.opened\",\"source\":\"panel\"}"}'
Invoke-RestMethod -Method Get -Uri http://localhost:3200/analytics/logs
```

Expected: logs include one event with `answer.eventName = panel.opened`.

- [ ] **Step 6: Inspect git status and touched high-conflict files**

Run:

```powershell
git status --short
git diff --stat
```

Expected: no unstaged changes except intentional files if verification updated snapshots. High-conflict files touched by this plan: none of `package.json`, lockfiles, CI, or installer definitions. The VSTO `.csproj` is touched because it is required for explicit compile inclusion.

- [ ] **Step 7: Stop on verification failures**

If any verification command fails, do not make a generic cleanup commit. Return to the task that owns the failing area, add or adjust the focused test first, then implement the smallest fix and commit it with that task's file list.
