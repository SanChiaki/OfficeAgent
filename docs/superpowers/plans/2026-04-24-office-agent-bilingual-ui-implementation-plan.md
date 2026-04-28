# OfficeAgent Bilingual UI Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add Excel-driven Chinese/English UI switching across the OfficeAgent task pane, Ribbon, host dialogs, and host-generated system messages while preserving AI free-form replies in the user’s input language.

**Architecture:** Resolve one host locale in the Excel add-in (`zh` or `en`) from `uiLanguageOverride` plus Excel UI language, expose it through a dedicated `bridge.getHostContext` payload, and localize host-generated strings and task-pane-generated strings independently. Keep persistence and bridge contracts backward-compatible so existing settings flows continue to work while reserving the `uiLanguageOverride` field for future manual language switching.

**Tech Stack:** C# (.NET Framework 4.8, VSTO, WebView2, Newtonsoft.Json), React 18, TypeScript, Vite, Vitest, xUnit

---

## File Structure

- Modify: `src/OfficeAgent.Core/Models/AppSettings.cs`
  Responsibility: add the persisted `uiLanguageOverride` field plus a shared normalization helper.
- Modify: `src/OfficeAgent.Infrastructure/Storage/FileSettingsStore.cs`
  Responsibility: load/save the override value without breaking existing settings files.
- Create: `src/OfficeAgent.ExcelAddIn/Localization/UiLocaleResolver.cs`
  Responsibility: convert `uiLanguageOverride + Excel UI language` into the supported `zh` / `en` locale.
- Create: `src/OfficeAgent.ExcelAddIn/Localization/HostLocalizedStrings.cs`
  Responsibility: centralize all host-generated Chinese/English UI strings.
- Modify: `src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj`
  Responsibility: explicitly include the new `Localization\*.cs` files because this VSTO project uses `Compile Include`.
- Modify: `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs`
  Responsibility: instantiate the locale resolver, translate Excel language IDs to culture names, expose the resolved locale, and construct `HostLocalizedStrings`.
- Modify: `src/OfficeAgent.ExcelAddIn/TaskPane/TaskPaneController.cs`
  Responsibility: pass the resolved-locale accessor and host strings into the task-pane host control.
- Modify: `src/OfficeAgent.ExcelAddIn/TaskPane/TaskPaneHostControl.cs`
  Responsibility: display localized WebView2 host errors when the task pane cannot initialize.
- Modify: `src/OfficeAgent.ExcelAddIn/WebBridge/WebMessageEnvelope.cs`
  Responsibility: declare `bridge.getHostContext` and the host-context payload type.
- Modify: `src/OfficeAgent.ExcelAddIn/WebBridge/WebMessageRouter.cs`
  Responsibility: return host context, accept the optional `uiLanguageOverride` settings field, and localize host-generated bridge errors.
- Modify: `src/OfficeAgent.ExcelAddIn/WebBridge/WebViewBootstrapper.cs`
  Responsibility: localize fallback HTML plus busy/timeout errors posted back to the task pane.
- Modify: `src/OfficeAgent.ExcelAddIn/AgentRibbon.Designer.cs`
  Responsibility: set English-safe default labels so design-time text never leaks Chinese into non-Chinese Excel before runtime refresh.
- Modify: `src/OfficeAgent.ExcelAddIn/AgentRibbon.cs`
  Responsibility: refresh all runtime Ribbon labels and status strings from `HostLocalizedStrings`.
- Modify: `src/OfficeAgent.ExcelAddIn/Dialogs/OperationResultDialog.cs`
  Responsibility: localize info/warning/error/authentication prompts.
- Modify: `src/OfficeAgent.ExcelAddIn/Dialogs/DownloadConfirmDialog.cs`
  Responsibility: localize the download confirmation body and title wrapper.
- Modify: `src/OfficeAgent.ExcelAddIn/Dialogs/UploadConfirmDialog.cs`
  Responsibility: localize the upload confirmation body and title wrapper.
- Modify: `src/OfficeAgent.ExcelAddIn/Dialogs/ProjectLayoutDialog.cs`
  Responsibility: localize dialog chrome, field instructions, and validation errors.
- Modify: `src/OfficeAgent.ExcelAddIn/SsoLoginPopup.cs`
  Responsibility: localize popup title and button captions.
- Modify: `src/OfficeAgent.Frontend/index.html`
  Responsibility: switch the default HTML `lang` to English so the task pane never flashes Chinese before host context resolves.
- Create: `src/OfficeAgent.Frontend/src/i18n/uiStrings.ts`
  Responsibility: hold all task-pane `zh` / `en` strings and helper formatters.
- Modify: `src/OfficeAgent.Frontend/src/types/bridge.ts`
  Responsibility: add `uiLanguageOverride` to `AppSettings` and declare the new `HostContext` bridge type.
- Modify: `src/OfficeAgent.Frontend/src/bridge/nativeBridge.ts`
  Responsibility: expose `getHostContext()`, round-trip `uiLanguageOverride`, and keep browser-preview defaults aligned with the new contract.
- Modify: `src/OfficeAgent.Frontend/src/components/ConfirmationCard.tsx`
  Responsibility: render localized labels passed from the app instead of hard-coded Chinese.
- Modify: `src/OfficeAgent.Frontend/src/App.tsx`
  Responsibility: fetch host locale, route all fixed UI/system messages through `uiStrings`, preserve `uiLanguageOverride` in settings saves, and localize session titles/welcome copy/plan formatting.
- Modify: `tests/OfficeAgent.Infrastructure.Tests/FileSettingsStoreTests.cs`
  Responsibility: lock persistence defaults and round-tripping for `uiLanguageOverride`.
- Create: `tests/OfficeAgent.ExcelAddIn.Tests/UiLocaleResolverTests.cs`
  Responsibility: lock `zh-*` versus non-`zh-*` resolution behavior.
- Modify: `tests/OfficeAgent.ExcelAddIn.Tests/WebMessageRouterTests.cs`
  Responsibility: lock the new `bridge.getHostContext` response and `uiLanguageOverride` settings persistence through the bridge.
- Create: `tests/OfficeAgent.ExcelAddIn.Tests/HostLocalizedStringsTests.cs`
  Responsibility: lock Chinese and English host strings without forcing UI reflection tests for every caption.
- Modify: `tests/OfficeAgent.ExcelAddIn.Tests/ProjectLayoutDialogTests.cs`
  Responsibility: verify the dialog renders English chrome when given an English host-string provider.
- Modify: `tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs`
  Responsibility: update source-level assertions to match the localized Ribbon implementation and English designer defaults.
- Modify: `src/OfficeAgent.Frontend/src/bridge/nativeBridge.test.ts`
  Responsibility: lock the new host-context bridge call plus preview-mode settings defaults.
- Modify: `src/OfficeAgent.Frontend/src/App.test.tsx`
  Responsibility: mock `getHostContext`, keep existing Chinese-path tests stable, and add English rendering coverage.
- Modify: `docs/modules/ribbon-sync-current-behavior.md`
  Responsibility: document bilingual Ribbon/login/dialog behavior.
- Create: `docs/modules/task-pane-current-behavior.md`
  Responsibility: snapshot task-pane bilingual behavior, system messages, and browser-preview defaults.
- Modify: `docs/module-index.md`
  Responsibility: register the new task-pane module snapshot.
- Modify: `docs/vsto-manual-test-checklist.md`
  Responsibility: add Chinese Excel / English Excel verification steps for the bilingual UI.

### Task 1: Persist `uiLanguageOverride` and add the locale resolver

**Files:**
- Modify: `src/OfficeAgent.Core/Models/AppSettings.cs`
- Modify: `src/OfficeAgent.Infrastructure/Storage/FileSettingsStore.cs`
- Create: `src/OfficeAgent.ExcelAddIn/Localization/UiLocaleResolver.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj`
- Modify: `tests/OfficeAgent.Infrastructure.Tests/FileSettingsStoreTests.cs`
- Create: `tests/OfficeAgent.ExcelAddIn.Tests/UiLocaleResolverTests.cs`

- [ ] **Step 1: Write the failing tests**

Add the persistence assertions to `tests/OfficeAgent.Infrastructure.Tests/FileSettingsStoreTests.cs`:

```csharp
[Fact]
public void LoadDefaultsUiLanguageOverrideToSystem()
{
    var store = new FileSettingsStore(Path.Combine(tempDirectory, "settings.json"), new DpapiSecretProtector());

    var settings = store.Load();

    Assert.Equal("system", settings.UiLanguageOverride);
}

[Fact]
public void SaveRoundTripsUiLanguageOverride()
{
    var settingsPath = Path.Combine(tempDirectory, "settings.json");
    var store = new FileSettingsStore(settingsPath, new DpapiSecretProtector());

    store.Save(new OfficeAgent.Core.Models.AppSettings
    {
        ApiKey = "secret-token",
        BaseUrl = "https://api.internal.example",
        BusinessBaseUrl = "https://business.internal.example",
        Model = "gpt-5-mini",
        UiLanguageOverride = "zh",
    });

    var loaded = store.Load();

    Assert.Equal("zh", loaded.UiLanguageOverride);
}
```

Create `tests/OfficeAgent.ExcelAddIn.Tests/UiLocaleResolverTests.cs`:

```csharp
using System;
using System.IO;
using System.Reflection;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class UiLocaleResolverTests
    {
        [Theory]
        [InlineData("system", "zh-CN", "zh")]
        [InlineData("system", "zh-TW", "zh")]
        [InlineData("system", "en-US", "en")]
        [InlineData("system", "fr-FR", "en")]
        [InlineData("system", "", "en")]
        public void ResolveMapsExcelUiLocaleToSupportedValues(string uiLanguageOverride, string excelUiLocale, string expected)
        {
            var resolverType = Assembly.LoadFrom(ResolveAddInAssemblyPath())
                .GetType("OfficeAgent.ExcelAddIn.Localization.UiLocaleResolver", throwOnError: true);
            var resolver = Activator.CreateInstance(
                resolverType,
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                binder: null,
                args: new object[] { (Func<string>)(() => excelUiLocale) },
                culture: null);
            var method = resolverType.GetMethod("Resolve", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

            Assert.Equal(expected, (string)method.Invoke(resolver, new object[] { uiLanguageOverride }));
        }

        [Theory]
        [InlineData("zh", "fr-FR", "zh")]
        [InlineData("en", "zh-CN", "en")]
        [InlineData("SYSTEM", "zh-CN", "zh")]
        [InlineData("garbage", "zh-CN", "zh")]
        public void ResolveHonorsExplicitOverrideAndNormalizesInvalidValues(string uiLanguageOverride, string excelUiLocale, string expected)
        {
            var resolverType = Assembly.LoadFrom(ResolveAddInAssemblyPath())
                .GetType("OfficeAgent.ExcelAddIn.Localization.UiLocaleResolver", throwOnError: true);
            var resolver = Activator.CreateInstance(
                resolverType,
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                binder: null,
                args: new object[] { (Func<string>)(() => excelUiLocale) },
                culture: null);
            var method = resolverType.GetMethod("Resolve", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

            Assert.Equal(expected, (string)method.Invoke(resolver, new object[] { uiLanguageOverride }));
        }

        private static string ResolveAddInAssemblyPath()
        {
            return Path.GetFullPath(
                Path.Combine(
                    AppContext.BaseDirectory,
                    "..",
                    "..",
                    "..",
                    "..",
                    "..",
                    "src",
                    "OfficeAgent.ExcelAddIn",
                    "bin",
                    "Debug",
                    "OfficeAgent.ExcelAddIn.dll"));
        }
    }
}
```

- [ ] **Step 2: Run the tests to verify they fail**

Run:

```powershell
dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj --filter FullyQualifiedName~FileSettingsStoreTests
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~UiLocaleResolverTests
```

Expected:

- `FileSettingsStoreTests` fails because `AppSettings` has no `UiLanguageOverride`.
- `UiLocaleResolverTests` fails because `OfficeAgent.ExcelAddIn.Localization.UiLocaleResolver` does not exist yet.

- [ ] **Step 3: Write the minimal implementation**

Update `src/OfficeAgent.Core/Models/AppSettings.cs`:

```csharp
public sealed class AppSettings
{
    public const string DefaultBaseUrl = "https://api.example.com";
    public const string DefaultUiLanguageOverride = "system";

    public string ApiKey { get; set; } = string.Empty;
    public string BaseUrl { get; set; } = DefaultBaseUrl;
    public string BusinessBaseUrl { get; set; } = string.Empty;
    public string Model { get; set; } = "gpt-5-mini";
    public string SsoUrl { get; set; } = string.Empty;
    public string SsoLoginSuccessPath { get; set; } = string.Empty;
    public string UiLanguageOverride { get; set; } = DefaultUiLanguageOverride;

    public static string NormalizeUiLanguageOverride(string value)
    {
        var normalized = (value ?? string.Empty).Trim().ToLowerInvariant();
        return normalized == "zh" || normalized == "en"
            ? normalized
            : DefaultUiLanguageOverride;
    }
}
```

Update `src/OfficeAgent.Infrastructure/Storage/FileSettingsStore.cs`:

```csharp
var settings = new AppSettings
{
    ApiKey = string.Empty,
    BaseUrl = AppSettings.NormalizeBaseUrl(persisted.BaseUrl),
    BusinessBaseUrl = AppSettings.NormalizeOptionalUrl(persisted.BusinessBaseUrl),
    Model = string.IsNullOrWhiteSpace(persisted.Model) ? "gpt-5-mini" : persisted.Model,
    SsoUrl = persisted.SsoUrl ?? string.Empty,
    SsoLoginSuccessPath = persisted.SsoLoginSuccessPath ?? string.Empty,
    UiLanguageOverride = AppSettings.NormalizeUiLanguageOverride(persisted.UiLanguageOverride),
};
```

```csharp
var persisted = new PersistedSettings
{
    EncryptedApiKey = secretProtector.Protect(settings?.ApiKey ?? string.Empty),
    BaseUrl = AppSettings.NormalizeBaseUrl(settings?.BaseUrl),
    BusinessBaseUrl = AppSettings.NormalizeOptionalUrl(settings?.BusinessBaseUrl),
    Model = string.IsNullOrWhiteSpace(settings?.Model) ? "gpt-5-mini" : settings.Model,
    SsoUrl = settings?.SsoUrl ?? string.Empty,
    SsoLoginSuccessPath = settings?.SsoLoginSuccessPath ?? string.Empty,
    UiLanguageOverride = AppSettings.NormalizeUiLanguageOverride(settings?.UiLanguageOverride),
};
```

```csharp
private sealed class PersistedSettings
{
    public string EncryptedApiKey { get; set; } = string.Empty;
    public string BaseUrl { get; set; } = string.Empty;
    public string BusinessBaseUrl { get; set; } = string.Empty;
    public string Model { get; set; } = string.Empty;
    public string SsoUrl { get; set; } = string.Empty;
    public string SsoLoginSuccessPath { get; set; } = string.Empty;
    public string UiLanguageOverride { get; set; } = AppSettings.DefaultUiLanguageOverride;
}
```

Create `src/OfficeAgent.ExcelAddIn/Localization/UiLocaleResolver.cs`:

```csharp
using System;
using OfficeAgent.Core.Models;

namespace OfficeAgent.ExcelAddIn.Localization
{
    internal sealed class UiLocaleResolver
    {
        private readonly Func<string> getExcelUiLocale;

        public UiLocaleResolver(Func<string> getExcelUiLocale)
        {
            this.getExcelUiLocale = getExcelUiLocale ?? throw new ArgumentNullException(nameof(getExcelUiLocale));
        }

        public string Resolve(string uiLanguageOverride)
        {
            var normalizedOverride = AppSettings.NormalizeUiLanguageOverride(uiLanguageOverride);
            if (!string.Equals(normalizedOverride, AppSettings.DefaultUiLanguageOverride, StringComparison.Ordinal))
            {
                return normalizedOverride;
            }

            var excelUiLocale = (getExcelUiLocale() ?? string.Empty).Trim();
            return excelUiLocale.StartsWith("zh", StringComparison.OrdinalIgnoreCase)
                ? "zh"
                : "en";
        }
    }
}
```

Register the new file in `src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj`:

```xml
<Compile Include="Localization\UiLocaleResolver.cs" />
```

- [ ] **Step 4: Run the tests to verify they pass**

Run:

```powershell
dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj --filter FullyQualifiedName~FileSettingsStoreTests
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~UiLocaleResolverTests
```

Expected: both commands PASS with `0` failed tests.

- [ ] **Step 5: Commit**

```powershell
git add src/OfficeAgent.Core/Models/AppSettings.cs src/OfficeAgent.Infrastructure/Storage/FileSettingsStore.cs src/OfficeAgent.ExcelAddIn/Localization/UiLocaleResolver.cs src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj tests/OfficeAgent.Infrastructure.Tests/FileSettingsStoreTests.cs tests/OfficeAgent.ExcelAddIn.Tests/UiLocaleResolverTests.cs
git commit -m "build: persist ui language override"
```

### Task 2: Expose the resolved host locale through the bridge

**Files:**
- Modify: `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/TaskPane/TaskPaneController.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/TaskPane/TaskPaneHostControl.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/WebBridge/WebMessageEnvelope.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/WebBridge/WebMessageRouter.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/WebBridge/WebViewBootstrapper.cs`
- Modify: `tests/OfficeAgent.ExcelAddIn.Tests/WebMessageRouterTests.cs`

- [ ] **Step 1: Write the failing bridge tests**

Add these tests to `tests/OfficeAgent.ExcelAddIn.Tests/WebMessageRouterTests.cs`:

```csharp
[Fact]
public void GetHostContextReturnsResolvedLocaleAndPersistedOverride()
{
    var sessionStore = new FileSessionStore(Path.Combine(tempDirectory, "sessions"));
    var settingsStore = new FileSettingsStore(
        Path.Combine(tempDirectory, "settings.json"),
        new DpapiSecretProtector());
    settingsStore.Save(new AppSettings
    {
        ApiKey = "secret-token",
        BaseUrl = "https://api.internal.example",
        BusinessBaseUrl = "https://business.internal.example",
        Model = "gpt-5-mini",
        UiLanguageOverride = "system",
    });

    var router = CreateRouter(sessionStore, settingsStore, resolvedUiLocale: "zh");
    var responseJson = InvokeRoute(router, "{\"type\":\"bridge.getHostContext\",\"requestId\":\"req-1\"}");

    Assert.Contains("\"ok\":true", responseJson);
    Assert.Contains("\"resolvedUiLocale\":\"zh\"", responseJson);
    Assert.Contains("\"uiLanguageOverride\":\"system\"", responseJson);
}

[Fact]
public void SaveSettingsRoundTripsUiLanguageOverride()
{
    var sessionStore = new FileSessionStore(Path.Combine(tempDirectory, "sessions"));
    var settingsStore = new FileSettingsStore(
        Path.Combine(tempDirectory, "settings.json"),
        new DpapiSecretProtector());

    var router = CreateRouter(sessionStore, settingsStore, resolvedUiLocale: "en");
    var responseJson = InvokeRoute(
        router,
        "{\"type\":\"bridge.saveSettings\",\"requestId\":\"req-1\",\"payload\":{\"apiKey\":\"secret-token\",\"baseUrl\":\"https://llm.internal.example\",\"businessBaseUrl\":\"https://business.internal.example\",\"model\":\"gpt-5-mini\",\"ssoUrl\":\"\",\"ssoLoginSuccessPath\":\"\",\"uiLanguageOverride\":\"en\"}}");

    Assert.Contains("\"ok\":true", responseJson);
    Assert.Contains("\"uiLanguageOverride\":\"en\"", responseJson);
    Assert.Equal("en", settingsStore.Load().UiLanguageOverride);
}
```

Also add a router helper overload that the implementation will satisfy:

```csharp
private static object CreateRouter(
    FileSessionStore sessionStore,
    FileSettingsStore settingsStore,
    string resolvedUiLocale)
{
    return CreateRouter(
        sessionStore,
        settingsStore,
        new FakeExcelContextService(SelectionContext.Empty("No selection available.")),
        new FakeExcelCommandExecutor(),
        new FakeAgentOrchestrator(),
        resolvedUiLocale);
}
```

- [ ] **Step 2: Run the Excel add-in tests to verify they fail**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~WebMessageRouterTests
```

Expected:

- the new tests fail because `bridge.getHostContext` is unknown,
- `WebMessageRouter` cannot round-trip `uiLanguageOverride`,
- and the new `CreateRouter(..., resolvedUiLocale)` helper does not match the current constructor signature.

- [ ] **Step 3: Write the minimal bridge implementation**

Add the bridge contract in `src/OfficeAgent.ExcelAddIn/WebBridge/WebMessageEnvelope.cs`:

```csharp
internal static class BridgeMessageTypes
{
    public const string Ping = "bridge.ping";
    public const string GetHostContext = "bridge.getHostContext";
    public const string GetSettings = "bridge.getSettings";
    public const string GetSelectionContext = "bridge.getSelectionContext";
    public const string SelectionContextChanged = "bridge.selectionContextChanged";
    public const string GetSessions = "bridge.getSessions";
    public const string SaveSessions = "bridge.saveSessions";
    public const string SaveSettings = "bridge.saveSettings";
    public const string ExecuteExcelCommand = "bridge.executeExcelCommand";
    public const string RunSkill = "bridge.runSkill";
    public const string RunAgent = "bridge.runAgent";
    public const string Login = "bridge.login";
    public const string Logout = "bridge.logout";
    public const string GetLoginStatus = "bridge.getLoginStatus";
}

internal sealed class HostContextPayload
{
    [JsonProperty("resolvedUiLocale")]
    public string ResolvedUiLocale { get; set; } = "en";

    [JsonProperty("uiLanguageOverride")]
    public string UiLanguageOverride { get; set; } = AppSettings.DefaultUiLanguageOverride;
}
```

Update `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs` to resolve the Excel UI language once per request:

```csharp
using System.Globalization;
using Microsoft.Office.Core;
using OfficeAgent.ExcelAddIn.Localization;

internal UiLocaleResolver UiLocaleResolver { get; private set; }

internal string GetResolvedUiLocale()
{
    var settings = SettingsStore?.Load() ?? new AppSettings();
    return UiLocaleResolver?.Resolve(settings.UiLanguageOverride) ?? "en";
}

private string GetExcelUiLocale()
{
    try
    {
        var languageId = Application?.LanguageSettings?.LanguageID[MsoAppLanguageID.msoLanguageIDUI] ?? 0;
        return languageId > 0 ? new CultureInfo(languageId).Name : string.Empty;
    }
    catch
    {
        return string.Empty;
    }
}
```

Initialize it during startup and thread the delegate into the task pane:

```csharp
UiLocaleResolver = new UiLocaleResolver(GetExcelUiLocale);
TaskPaneController = new TaskPaneController(
    this,
    SessionStore,
    SettingsStore,
    ExcelContextService,
    ExcelCommandExecutor,
    AgentOrchestrator,
    SharedCookies,
    CookieStore,
    GetResolvedUiLocale);
```

Thread the delegate through `TaskPaneController`, `TaskPaneHostControl`, and `WebViewBootstrapper`:

```csharp
public TaskPaneController(
    ThisAddIn addIn,
    FileSessionStore sessionStore,
    FileSettingsStore settingsStore,
    IExcelContextService excelContextService,
    IExcelCommandExecutor excelCommandExecutor,
    IAgentOrchestrator agentOrchestrator,
    SharedCookieContainer sharedCookies,
    FileCookieStore cookieStore,
    Func<string> getResolvedUiLocale)
{
    // existing assignments...
    this.getResolvedUiLocale = getResolvedUiLocale ?? throw new ArgumentNullException(nameof(getResolvedUiLocale));
}
```

Update `src/OfficeAgent.ExcelAddIn/WebBridge/WebMessageRouter.cs`:

```csharp
private readonly Func<string> getResolvedUiLocale;

public WebMessageRouter(
    FileSessionStore sessionStore,
    FileSettingsStore settingsStore,
    IExcelContextService excelContextService,
    IExcelCommandExecutor excelCommandExecutor,
    IAgentOrchestrator agentOrchestrator,
    SharedCookieContainer sharedCookies,
    FileCookieStore cookieStore,
    Func<string> getResolvedUiLocale)
{
    // existing assignments...
    this.getResolvedUiLocale = getResolvedUiLocale ?? throw new ArgumentNullException(nameof(getResolvedUiLocale));
}
```

```csharp
private readonly HashSet<string> allowedTypes = new HashSet<string>(StringComparer.Ordinal)
{
    BridgeMessageTypes.Ping,
    BridgeMessageTypes.GetHostContext,
    BridgeMessageTypes.GetSettings,
    // existing items...
};
```

```csharp
case BridgeMessageTypes.GetHostContext:
    if (HasUnexpectedPayload(request.Payload))
    {
        return Error(
            request.Type,
            request.RequestId,
            code: "malformed_payload",
            message: "bridge.getHostContext does not accept a payload.");
    }

    var hostContextSettings = settingsStore.Load();
    return Success(
        request.Type,
        request.RequestId,
        new HostContextPayload
        {
            ResolvedUiLocale = getResolvedUiLocale(),
            UiLanguageOverride = AppSettings.NormalizeUiLanguageOverride(hostContextSettings.UiLanguageOverride),
        });
```

Make `uiLanguageOverride` optional-but-preserved in `HasValidSettingsPayload`:

```csharp
var uiLanguageOverrideToken = payloadObject["uiLanguageOverride"];
return IsStringToken(payloadObject["apiKey"]) &&
       IsStringToken(payloadObject["baseUrl"]) &&
       IsStringToken(payloadObject["businessBaseUrl"]) &&
       IsStringToken(payloadObject["model"]) &&
       (uiLanguageOverrideToken == null || IsStringToken(uiLanguageOverrideToken)) &&
       payloadObject.Count >= 4;
```

Update the `WebMessageRouterTests` reflection helper so it passes the locale delegate:

```csharp
args: new object[] { sessionStore, settingsStore, selectionContextService, excelCommandExecutor, agentOrchestrator, sharedCookies, cookieStore, (Func<string>)(() => resolvedUiLocale) },
```

- [ ] **Step 4: Run the bridge tests to verify they pass**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~WebMessageRouterTests
```

Expected: PASS with `0` failed tests.

- [ ] **Step 5: Commit**

```powershell
git add src/OfficeAgent.ExcelAddIn/ThisAddIn.cs src/OfficeAgent.ExcelAddIn/TaskPane/TaskPaneController.cs src/OfficeAgent.ExcelAddIn/TaskPane/TaskPaneHostControl.cs src/OfficeAgent.ExcelAddIn/WebBridge/WebMessageEnvelope.cs src/OfficeAgent.ExcelAddIn/WebBridge/WebMessageRouter.cs src/OfficeAgent.ExcelAddIn/WebBridge/WebViewBootstrapper.cs tests/OfficeAgent.ExcelAddIn.Tests/WebMessageRouterTests.cs
git commit -m "feat: expose host locale through bridge"
```

### Task 3: Add the frontend host-context contract and preserve the override in preview settings

**Files:**
- Modify: `src/OfficeAgent.Frontend/src/types/bridge.ts`
- Modify: `src/OfficeAgent.Frontend/src/bridge/nativeBridge.ts`
- Modify: `src/OfficeAgent.Frontend/src/bridge/nativeBridge.test.ts`

- [ ] **Step 1: Write the failing frontend bridge tests**

Add to `src/OfficeAgent.Frontend/src/bridge/nativeBridge.test.ts`:

```ts
it('returns host context in browser preview mode', async () => {
  const bridge = new NativeBridge(undefined);

  await expect(bridge.getHostContext()).resolves.toEqual({
    resolvedUiLocale: 'en',
    uiLanguageOverride: 'system',
  });
});

it('sends getHostContext requests through the structured bridge contract', async () => {
  const webView = createMockWebView();
  const bridge = new NativeBridge(webView);

  const pending = bridge.getHostContext();
  const [request] = webView.postedMessages as Array<{ type: string; requestId: string }>;

  expect(request.type).toBe('bridge.getHostContext');

  webView.dispatch({
    type: 'bridge.getHostContext',
    requestId: request.requestId,
    ok: true,
    payload: {
      resolvedUiLocale: 'zh',
      uiLanguageOverride: 'system',
    },
  });

  await expect(pending).resolves.toEqual({
    resolvedUiLocale: 'zh',
    uiLanguageOverride: 'system',
  });
});

it('includes uiLanguageOverride in browser preview settings', async () => {
  const bridge = new NativeBridge(undefined);

  await expect(bridge.getSettings()).resolves.toMatchObject({
    uiLanguageOverride: 'system',
  });
});
```

- [ ] **Step 2: Run the Vitest file to verify it fails**

Run:

```powershell
cd src/OfficeAgent.Frontend
npm run test -- src/bridge/nativeBridge.test.ts
```

Expected: FAIL because `HostContext` and `getHostContext()` do not exist yet, and preview settings do not carry `uiLanguageOverride`.

- [ ] **Step 3: Write the minimal frontend bridge implementation**

Update `src/OfficeAgent.Frontend/src/types/bridge.ts`:

```ts
export interface AppSettings {
  apiKey: string;
  baseUrl: string;
  businessBaseUrl: string;
  model: string;
  ssoUrl: string;
  ssoLoginSuccessPath: string;
  uiLanguageOverride: 'system' | 'zh' | 'en';
}

export interface HostContext {
  resolvedUiLocale: 'zh' | 'en';
  uiLanguageOverride: 'system' | 'zh' | 'en';
}
```

Update `src/OfficeAgent.Frontend/src/bridge/nativeBridge.ts`:

```ts
import type {
  AppSettings,
  HostContext,
  // existing imports...
} from '../types/bridge';

const BRIDGE_TYPES = {
  ping: 'bridge.ping',
  getHostContext: 'bridge.getHostContext',
  getSettings: 'bridge.getSettings',
  // existing bridge types...
} as const;

const BROWSER_PREVIEW_HOST_CONTEXT: HostContext = {
  resolvedUiLocale: 'en',
  uiLanguageOverride: 'system',
};

const BROWSER_PREVIEW_SETTINGS: AppSettings = {
  apiKey: '',
  baseUrl: 'https://api.example.com',
  businessBaseUrl: '',
  model: 'gpt-5-mini',
  ssoUrl: '',
  ssoLoginSuccessPath: '',
  uiLanguageOverride: 'system',
};
```

```ts
getHostContext() {
  return this.invoke<void, HostContext>(BRIDGE_TYPES.getHostContext);
}
```

```ts
if (type === BRIDGE_TYPES.getHostContext) {
  return Promise.resolve(BROWSER_PREVIEW_HOST_CONTEXT as TResult);
}
```

Keep the preview `saveSettings` shape round-trippable:

```ts
uiLanguageOverride: typeof (payload as AppSettings | undefined)?.uiLanguageOverride === 'string'
  ? (payload as AppSettings).uiLanguageOverride
  : BROWSER_PREVIEW_SETTINGS.uiLanguageOverride,
```

- [ ] **Step 4: Run the Vitest file to verify it passes**

Run:

```powershell
cd src/OfficeAgent.Frontend
npm run test -- src/bridge/nativeBridge.test.ts
```

Expected: PASS with `0` failed tests.

- [ ] **Step 5: Commit**

```powershell
git add src/OfficeAgent.Frontend/src/types/bridge.ts src/OfficeAgent.Frontend/src/bridge/nativeBridge.ts src/OfficeAgent.Frontend/src/bridge/nativeBridge.test.ts
git commit -m "feat: add frontend host context bridge contract"
```

### Task 4: Centralize host-localized strings and wire them into Ribbon, dialogs, and host error chrome

**Files:**
- Create: `src/OfficeAgent.ExcelAddIn/Localization/HostLocalizedStrings.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj`
- Modify: `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/AgentRibbon.Designer.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/AgentRibbon.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/Dialogs/OperationResultDialog.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/Dialogs/DownloadConfirmDialog.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/Dialogs/UploadConfirmDialog.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/Dialogs/ProjectLayoutDialog.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/SsoLoginPopup.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/TaskPane/TaskPaneHostControl.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/WebBridge/WebMessageRouter.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/WebBridge/WebViewBootstrapper.cs`
- Create: `tests/OfficeAgent.ExcelAddIn.Tests/HostLocalizedStringsTests.cs`
- Modify: `tests/OfficeAgent.ExcelAddIn.Tests/ProjectLayoutDialogTests.cs`
- Modify: `tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs`

- [ ] **Step 1: Write the failing host-localization tests**

Create `tests/OfficeAgent.ExcelAddIn.Tests/HostLocalizedStringsTests.cs`:

```csharp
using System;
using System.IO;
using System.Reflection;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class HostLocalizedStringsTests
    {
        [Fact]
        public void ChineseLocaleReturnsChineseHostStrings()
        {
            var stringsType = Assembly.LoadFrom(ResolveAddInAssemblyPath())
                .GetType("OfficeAgent.ExcelAddIn.Localization.HostLocalizedStrings", throwOnError: true);
            var strings = Activator.CreateInstance(
                stringsType,
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                binder: null,
                args: new object[] { (Func<string>)(() => "zh") },
                culture: null);

            Assert.Equal("先选择项目", Invoke(stringsType, strings, "SelectProjectPlaceholder"));
            Assert.Equal("请先登录", Invoke(stringsType, strings, "LoginRequiredStatus"));
            Assert.Equal("配置当前表布局", Invoke(stringsType, strings, "ProjectLayoutDialogTitle"));
        }

        [Fact]
        public void EnglishLocaleReturnsEnglishHostStrings()
        {
            var stringsType = Assembly.LoadFrom(ResolveAddInAssemblyPath())
                .GetType("OfficeAgent.ExcelAddIn.Localization.HostLocalizedStrings", throwOnError: true);
            var strings = Activator.CreateInstance(
                stringsType,
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                binder: null,
                args: new object[] { (Func<string>)(() => "en") },
                culture: null);

            Assert.Equal("Select a project", Invoke(stringsType, strings, "SelectProjectPlaceholder"));
            Assert.Equal("Sign in required", Invoke(stringsType, strings, "LoginRequiredStatus"));
            Assert.Equal("Configure current sheet layout", Invoke(stringsType, strings, "ProjectLayoutDialogTitle"));
            Assert.Equal("WebView2 Runtime is required to render ISDP.", Invoke(stringsType, strings, "WebViewRuntimeMissing"));
        }

        private static string Invoke(Type stringsType, object instance, string methodName)
        {
            return (string)stringsType
                .GetMethod(methodName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
                .Invoke(instance, null);
        }

        private static string ResolveAddInAssemblyPath()
        {
            return Path.GetFullPath(
                Path.Combine(
                    AppContext.BaseDirectory,
                    "..",
                    "..",
                    "..",
                    "..",
                    "..",
                    "src",
                    "OfficeAgent.ExcelAddIn",
                    "bin",
                    "Debug",
                    "OfficeAgent.ExcelAddIn.dll"));
        }
    }
}
```

Update `tests/OfficeAgent.ExcelAddIn.Tests/ProjectLayoutDialogTests.cs` to prove English chrome is applied:

```csharp
[Fact]
public void EnglishDialogUsesLocalizedChrome()
{
    RunInSta(() =>
    {
        using (var dialog = CreateDialog("en"))
        {
            dialog.CreateControl();

            Assert.Equal("Configure current sheet layout", dialog.Text);
            Assert.Contains(dialog.Controls.Find("HeaderStartRowTextBox", true), control => control is TextBox);
            Assert.Contains(dialog.Controls.Cast<Control>().SelectMany(EnumerateParents), control => control is Button button && button.Text == "Save");
            Assert.Contains(dialog.Controls.Cast<Control>().SelectMany(EnumerateParents), control => control is Button button && button.Text == "Cancel");
        }
    });
}

private static Form CreateDialog(string locale)
{
    var addInAssembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
    var stringsType = addInAssembly.GetType("OfficeAgent.ExcelAddIn.Localization.HostLocalizedStrings", throwOnError: true);
    var strings = Activator.CreateInstance(
        stringsType,
        BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
        binder: null,
        args: new object[] { (Func<string>)(() => locale) },
        culture: null);

    return (Form)Activator.CreateInstance(
        GetProjectLayoutDialogType(),
        BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
        binder: null,
        args: new object[] { CreateSeedBinding(), strings },
        culture: null);
}
```

Update `tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs` assertions so they match the localized implementation:

```csharp
Assert.Contains("this.projectDropDown.Label = \"Select a project\";", designerText, StringComparison.Ordinal);
Assert.Contains("hostStrings.SelectProjectPlaceholder()", ribbonCodeText, StringComparison.Ordinal);
Assert.Contains("hostStrings.LoginRequiredStatus()", ribbonCodeText, StringComparison.Ordinal);
Assert.Contains("hostStrings.ProjectLoadFailedStatus()", ribbonCodeText, StringComparison.Ordinal);
Assert.Contains("ApplyLocalizedLabels();", ribbonCodeText, StringComparison.Ordinal);
```

- [ ] **Step 2: Run the Excel add-in tests to verify they fail**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj
```

Expected: FAIL in the new host-localization tests and in any source-text assertions that still expect hard-coded Chinese literals.

- [ ] **Step 3: Write the minimal host-localization implementation**

Create `src/OfficeAgent.ExcelAddIn/Localization/HostLocalizedStrings.cs`:

```csharp
using System;

namespace OfficeAgent.ExcelAddIn.Localization
{
    internal sealed class HostLocalizedStrings
    {
        private readonly Func<string> getResolvedUiLocale;

        public HostLocalizedStrings(Func<string> getResolvedUiLocale)
        {
            this.getResolvedUiLocale = getResolvedUiLocale ?? throw new ArgumentNullException(nameof(getResolvedUiLocale));
        }

        private bool IsChinese => string.Equals(getResolvedUiLocale(), "zh", StringComparison.Ordinal);

        public string SelectProjectPlaceholder() => IsChinese ? "先选择项目" : "Select a project";
        public string LoginRequiredStatus() => IsChinese ? "请先登录" : "Sign in required";
        public string NoProjectsStatus() => IsChinese ? "无可用项目" : "No projects";
        public string ProjectLoadFailedStatus() => IsChinese ? "项目加载失败" : "Project load failed";
        public string LoggingInLabel() => IsChinese ? "登录中..." : "Signing in...";
        public string LoginLabel() => IsChinese ? "登录" : "Log in";
        public string ProjectGroupLabel() => IsChinese ? "项目" : "Project";
        public string DownloadGroupLabel() => IsChinese ? "下载" : "Download";
        public string UploadGroupLabel() => IsChinese ? "上传" : "Upload";
        public string AccountGroupLabel() => IsChinese ? "账号" : "Account";
        public string InitializeSheetLabel() => IsChinese ? "初始化当前表" : "Initialize sheet";
        public string PartialDownloadLabel() => IsChinese ? "部分下载" : "Partial download";
        public string PartialUploadLabel() => IsChinese ? "部分上传" : "Partial upload";
        public string ProjectLayoutDialogTitle() => IsChinese ? "配置当前表布局" : "Configure current sheet layout";
        public string ProjectLayoutInstruction() => IsChinese ? "下面三个值会写入当前工作表的同步配置（SheetBindings），请确认后保存。" : "These three values will be saved into the current worksheet sync binding (SheetBindings). Review them before saving.";
        public string SaveLabel() => IsChinese ? "保存" : "Save";
        public string CancelLabel() => IsChinese ? "取消" : "Cancel";
        public string SignedInLabel() => IsChinese ? "已登录" : "Signed in";
        public string WebViewRuntimeMissing() => IsChinese ? "渲染 ISDP 需要 WebView2 Runtime。" : "WebView2 Runtime is required to render ISDP.";
        public string TaskPaneInitializationFailed() => IsChinese ? "ISDP 无法初始化任务窗格。请检查本地日志并重新打开 Excel。" : "ISDP could not initialize the task pane. Check the local log and reopen Excel.";
        public string BridgeBusy() => IsChinese ? "当前已有请求正在执行，请稍后再试。" : "Another request is already in progress. Please wait.";
        public string BridgeTimedOut() => IsChinese ? "Agent 请求超时。" : "Agent request timed out.";
        public string BridgeUnexpectedError() => IsChinese ? "ISDP 遇到了意外错误。请检查本地日志后重试。" : "ISDP hit an unexpected error. Check the local log and try again.";
    }
}
```

Register the new file in `src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj`:

```xml
<Compile Include="Localization\HostLocalizedStrings.cs" />
```

Expose the host strings in `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs`:

```csharp
internal HostLocalizedStrings HostLocalizedStrings { get; private set; }

HostLocalizedStrings = new HostLocalizedStrings(GetResolvedUiLocale);
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
    HostLocalizedStrings);
```

Update the Ribbon runtime labels in `src/OfficeAgent.ExcelAddIn/AgentRibbon.cs`:

```csharp
private HostLocalizedStrings hostStrings => Globals.ThisAddIn.HostLocalizedStrings;

private string ProjectDropDownPlaceholderText => hostStrings.SelectProjectPlaceholder();

private string[] StickyNoProjectTexts => new[]
{
    hostStrings.LoginRequiredStatus(),
    hostStrings.NoProjectsStatus(),
    hostStrings.ProjectLoadFailedStatus(),
};

private void AgentRibbon_Load(object sender, RibbonUIEventArgs e)
{
    ApplyLocalizedLabels();
    SetProjectDropDownText(ProjectDropDownPlaceholderText);
    if (!TryBindToSyncController())
    {
        return;
    }

    var syncController = Globals.ThisAddIn.RibbonSyncController;
    if (syncController == null)
    {
        return;
    }

    syncController.RefreshActiveProjectFromSheetMetadata();
    RefreshProjectDropDownFromController();
}

private void ApplyLocalizedLabels()
{
    groupProject.Label = hostStrings.ProjectGroupLabel();
    groupDownload.Label = hostStrings.DownloadGroupLabel();
    groupUpload.Label = hostStrings.UploadGroupLabel();
    group2.Label = hostStrings.AccountGroupLabel();
    projectDropDown.Label = hostStrings.SelectProjectPlaceholder();
    initializeSheetButton.Label = hostStrings.InitializeSheetLabel();
    partialDownloadButton.Label = hostStrings.PartialDownloadLabel();
    partialUploadButton.Label = hostStrings.PartialUploadLabel();
    loginButton.Label = hostStrings.LoginLabel();
    RibbonUI?.Invalidate();
}
```

Set English-safe designer defaults in `src/OfficeAgent.ExcelAddIn/AgentRibbon.Designer.cs`:

```csharp
this.groupProject.Label = "Project";
this.projectDropDown.Label = "Select a project";
this.initializeSheetButton.Label = "Initialize sheet";
this.groupDownload.Label = "Download";
this.partialDownloadButton.Label = "Partial download";
this.groupUpload.Label = "Upload";
this.partialUploadButton.Label = "Partial upload";
this.group2.Label = "Account";
this.loginButton.Label = "Log in";
```

Pass `HostLocalizedStrings` through the dialog services:

```csharp
internal sealed class RibbonSyncDialogService : IRibbonSyncDialogService
{
    private readonly HostLocalizedStrings strings;

    public RibbonSyncDialogService(HostLocalizedStrings strings)
    {
        this.strings = strings ?? throw new ArgumentNullException(nameof(strings));
    }

    public SheetBinding ShowProjectLayoutDialog(SheetBinding suggestedBinding)
    {
        using (var dialog = new ProjectLayoutDialog(suggestedBinding, strings))
        {
            return dialog.ShowDialog() == DialogResult.OK
                ? dialog.ResultBinding
                : null;
        }
    }
}
```

Update `ProjectLayoutDialog`, `DownloadConfirmDialog`, `UploadConfirmDialog`, `OperationResultDialog`, and `SsoLoginPopup` to consume `HostLocalizedStrings` instead of hard-coded Chinese labels. Keep backend-provided values such as `projectName` or `ex.Message` verbatim; only localize the wrapper copy.

Localize the task-pane host fallback labels in `TaskPaneHostControl.cs`:

```csharp
Text = hostLocalizedStrings.WebViewRuntimeMissing(),
```

```csharp
Text = hostLocalizedStrings.TaskPaneInitializationFailed(),
```

Localize `WebViewBootstrapper` busy/time-out/fallback HTML messages:

```csharp
TryPostErrorResponse(rawJson, "busy", hostLocalizedStrings.BridgeBusy());
```

```csharp
var message = error is OperationCanceledException
    ? hostLocalizedStrings.BridgeTimedOut()
    : (error.Message ?? hostLocalizedStrings.BridgeUnexpectedError());
```

Localize the generic bridge error strings in `WebMessageRouter.cs` using `hostLocalizedStrings.BridgeUnexpectedError()` instead of hard-coded English or Chinese fallbacks.

- [ ] **Step 4: Run the Excel add-in tests to verify they pass**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj
```

Expected: PASS with `0` failed tests.

- [ ] **Step 5: Commit**

```powershell
git add src/OfficeAgent.ExcelAddIn/Localization/HostLocalizedStrings.cs src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj src/OfficeAgent.ExcelAddIn/ThisAddIn.cs src/OfficeAgent.ExcelAddIn/AgentRibbon.Designer.cs src/OfficeAgent.ExcelAddIn/AgentRibbon.cs src/OfficeAgent.ExcelAddIn/Dialogs/OperationResultDialog.cs src/OfficeAgent.ExcelAddIn/Dialogs/DownloadConfirmDialog.cs src/OfficeAgent.ExcelAddIn/Dialogs/UploadConfirmDialog.cs src/OfficeAgent.ExcelAddIn/Dialogs/ProjectLayoutDialog.cs src/OfficeAgent.ExcelAddIn/SsoLoginPopup.cs src/OfficeAgent.ExcelAddIn/TaskPane/TaskPaneHostControl.cs src/OfficeAgent.ExcelAddIn/WebBridge/WebMessageRouter.cs src/OfficeAgent.ExcelAddIn/WebBridge/WebViewBootstrapper.cs tests/OfficeAgent.ExcelAddIn.Tests/HostLocalizedStringsTests.cs tests/OfficeAgent.ExcelAddIn.Tests/ProjectLayoutDialogTests.cs tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs
git commit -m "feat: localize add-in host chrome"
```

### Task 5: Localize the React task pane and its generated system messages

**Files:**
- Modify: `src/OfficeAgent.Frontend/index.html`
- Create: `src/OfficeAgent.Frontend/src/i18n/uiStrings.ts`
- Modify: `src/OfficeAgent.Frontend/src/types/bridge.ts`
- Modify: `src/OfficeAgent.Frontend/src/bridge/nativeBridge.ts`
- Modify: `src/OfficeAgent.Frontend/src/components/ConfirmationCard.tsx`
- Modify: `src/OfficeAgent.Frontend/src/App.tsx`
- Modify: `src/OfficeAgent.Frontend/src/App.test.tsx`
- Modify: `src/OfficeAgent.Frontend/src/bridge/nativeBridge.test.ts`

- [ ] **Step 1: Write the failing React tests**

Update the mocked bridge in `src/OfficeAgent.Frontend/src/App.test.tsx` so the app can ask for host context:

```ts
vi.mock('./bridge/nativeBridge', () => ({
  nativeBridge: {
    ping: vi.fn(),
    getHostContext: vi.fn(),
    getSelectionContext: vi.fn(),
    getSessions: vi.fn(),
    onSelectionContextChanged: vi.fn(),
    getSettings: vi.fn(),
    saveSettings: vi.fn(),
    executeExcelCommand: vi.fn(),
    runSkill: vi.fn(),
    runAgent: vi.fn(),
    login: vi.fn(),
    logout: vi.fn(),
    getLoginStatus: vi.fn(),
  },
}));
```

Set the default mocked locale to Chinese so the existing Chinese-path tests remain stable:

```ts
mockedBridge.getHostContext.mockResolvedValue({
  resolvedUiLocale: 'zh',
  uiLanguageOverride: 'system',
});

mockedBridge.getSettings.mockResolvedValue({
  apiKey: '',
  baseUrl: 'https://api.example.com',
  businessBaseUrl: 'https://business.example.com',
  model: 'gpt-5-mini',
  ssoUrl: '',
  ssoLoginSuccessPath: '',
  uiLanguageOverride: 'system',
});
```

Add an explicit English rendering test:

```ts
it('renders English fixed UI when the host locale is en', async () => {
  mockedBridge.getHostContext.mockResolvedValueOnce({
    resolvedUiLocale: 'en',
    uiLanguageOverride: 'system',
  });

  render(<App />);

  expect(await screen.findByText(/welcome to isdp/i)).toBeInTheDocument();
  expect(screen.getByRole('button', { name: /open settings/i })).toBeInTheDocument();
  expect(screen.getByRole('button', { name: /open sessions/i })).toBeInTheDocument();
  expect(screen.getByRole('button', { name: /send/i })).toBeInTheDocument();
  expect(screen.getByRole('heading', { name: /new chat/i })).toBeInTheDocument();
});
```

Update the existing Chinese-path expectations so the untitled thread title becomes localized:

```ts
expect(await screen.findByRole('heading', { name: /新建会话/i })).toBeInTheDocument();
```

Add a preview-locale test to `src/OfficeAgent.Frontend/src/bridge/nativeBridge.test.ts`:

```ts
it('returns Chinese browser preview text when the preview locale is zh', async () => {
  const bridge = new NativeBridge(undefined, 'zh');

  await expect(bridge.getHostContext()).resolves.toEqual({
    resolvedUiLocale: 'zh',
    uiLanguageOverride: 'system',
  });
});
```

- [ ] **Step 2: Run the React tests to verify they fail**

Run:

```powershell
cd src/OfficeAgent.Frontend
npm run test -- src/App.test.tsx src/bridge/nativeBridge.test.ts
```

Expected: FAIL because `App` does not call `getHostContext`, the heading/title strings are still hard-coded, and the browser-preview constructor does not accept a locale override.

- [ ] **Step 3: Write the minimal frontend localization implementation**

Switch `src/OfficeAgent.Frontend/index.html` to an English-safe default:

```html
<html lang="en">
```

Create `src/OfficeAgent.Frontend/src/i18n/uiStrings.ts`:

```ts
export type UiLocale = 'zh' | 'en';

export type UiStrings = {
  newChatTitle: string;
  welcomeMessage: string;
  openSessions: string;
  closeSessions: string;
  openSettings: string;
  send: string;
  composerPlaceholder: string;
  settingsDialogLabel: string;
  settingsTitle: string;
  settingsEyebrow: string;
  close: string;
  cancel: string;
  save: string;
  login: string;
  loggingIn: string;
  logout: string;
  loggedIn: string;
  loggedOut: string;
  noSelection: string;
  sessionsTitle: string;
  noSessions: string;
  renameSession: string;
  confirmRename: string;
  cancelRename: string;
  deleteSession: string;
  deleteSessionTitle: string;
  confirmationAriaLabel: string;
  confirmationEyebrow: string;
  confirmationTitle: string;
  confirm: string;
  executePlanTitle: string;
  uploadSelectedData: string;
  bridgeConnecting: string;
  bridgeConnected: (host: string, version: string) => string;
  bridgeUnavailable: (message: string) => string;
  requestFailed: (message: string) => string;
  settingsLoadFailed: string;
  settingsSaveFailed: string;
  loginFailed: string;
  excelCommandFailed: string;
  skillFailed: string;
  agentFailed: string;
  cancelPendingUpload: string;
  cancelPendingPlan: string;
  cancelPendingExcel: string;
  deleteSessionPrompt: (title: string) => string;
  addWorksheetStep: (sheetName: string) => string;
  writeRangeStep: (targetAddress: string) => string;
  renameWorksheetStep: (sheetName: string, newSheetName: string) => string;
  deleteWorksheetStep: (sheetName: string) => string;
};

const zh: UiStrings = {
  newChatTitle: '新建会话',
  welcomeMessage: '欢迎使用ISDP，我是能和Excel交互的Agent。你选中的单元格会被我优先识别，尽情尝试吧~',
  openSessions: '打开会话列表',
  closeSessions: '关闭会话列表',
  openSettings: '打开设置',
  send: '发送',
  composerPlaceholder: '输入消息...',
  settingsDialogLabel: '设置对话框',
  settingsTitle: '设置',
  settingsEyebrow: '配置',
  close: '关闭',
  cancel: '取消',
  save: '保存',
  login: '登录',
  loggingIn: '登录中...',
  logout: '登出',
  loggedIn: '已登录',
  loggedOut: '未登录',
  noSelection: '未选中',
  sessionsTitle: '会话',
  noSessions: '暂无会话',
  renameSession: '重命名会话',
  confirmRename: '确认重命名',
  cancelRename: '取消重命名',
  deleteSession: '删除会话',
  deleteSessionTitle: '删除会话',
  confirmationAriaLabel: '确认 Excel 操作',
  confirmationEyebrow: '待确认的写入操作',
  confirmationTitle: '确认 Excel 操作',
  confirm: '确认',
  executePlanTitle: '执行计划',
  uploadSelectedData: '上传所选数据',
  bridgeConnecting: '正在连接宿主...',
  bridgeConnected: (host, version) => `已连接 ${host} (${version})`,
  bridgeUnavailable: (message) => `宿主不可用: ${message}`,
  requestFailed: (message) => `请求失败：${message}`,
  settingsLoadFailed: '无法从宿主加载设置。',
  settingsSaveFailed: '保存设置失败。',
  loginFailed: '登录失败。',
  excelCommandFailed: 'Excel 命令执行失败。',
  skillFailed: 'Skill 执行失败。',
  agentFailed: 'Agent 执行失败。',
  cancelPendingUpload: '已取消待处理的上传操作。',
  cancelPendingPlan: '已取消待执行的计划。',
  cancelPendingExcel: '已取消待处理的 Excel 操作。',
  deleteSessionPrompt: (title) => `确定要删除「${title}」吗？此操作不可撤销。`,
  addWorksheetStep: (sheetName) => `新增工作表 ${sheetName}`.trim(),
  writeRangeStep: (targetAddress) => `写入范围 ${targetAddress}`.trim(),
  renameWorksheetStep: (sheetName, newSheetName) => `重命名工作表 ${sheetName} 为 ${newSheetName}`.trim(),
  deleteWorksheetStep: (sheetName) => `删除工作表 ${sheetName}`.trim(),
};

const en: UiStrings = {
  newChatTitle: 'New chat',
  welcomeMessage: 'Welcome to ISDP. I can work with Excel selections, sheets, and guided actions inside the task pane.',
  openSessions: 'Open sessions',
  closeSessions: 'Close sessions',
  openSettings: 'Open settings',
  send: 'Send',
  composerPlaceholder: 'Type a message...',
  settingsDialogLabel: 'Settings dialog',
  settingsTitle: 'Settings',
  settingsEyebrow: 'Configuration',
  close: 'Close',
  cancel: 'Cancel',
  save: 'Save',
  login: 'Log in',
  loggingIn: 'Signing in...',
  logout: 'Log out',
  loggedIn: 'Signed in',
  loggedOut: 'Signed out',
  noSelection: 'No selection',
  sessionsTitle: 'Sessions',
  noSessions: 'No sessions yet',
  renameSession: 'Rename session',
  confirmRename: 'Confirm rename',
  cancelRename: 'Cancel rename',
  deleteSession: 'Delete session',
  deleteSessionTitle: 'Delete session',
  confirmationAriaLabel: 'Confirm Excel action',
  confirmationEyebrow: 'Pending workbook change',
  confirmationTitle: 'Confirm Excel action',
  confirm: 'Confirm',
  executePlanTitle: 'Execution plan',
  uploadSelectedData: 'Upload selected data',
  bridgeConnecting: 'Connecting to host...',
  bridgeConnected: (host, version) => `Connected to ${host} (${version})`,
  bridgeUnavailable: (message) => `Host unavailable: ${message}`,
  requestFailed: (message) => `Request failed: ${message}`,
  settingsLoadFailed: 'Unable to load settings from the host.',
  settingsSaveFailed: 'Failed to save settings.',
  loginFailed: 'Sign-in failed.',
  excelCommandFailed: 'Excel command execution failed.',
  skillFailed: 'Skill execution failed.',
  agentFailed: 'Agent execution failed.',
  cancelPendingUpload: 'Cancelled the pending upload.',
  cancelPendingPlan: 'Cancelled the pending plan.',
  cancelPendingExcel: 'Cancelled the pending Excel action.',
  deleteSessionPrompt: (title) => `Delete "${title}"? This action cannot be undone.`,
  addWorksheetStep: (sheetName) => `Add worksheet ${sheetName}`.trim(),
  writeRangeStep: (targetAddress) => `Write range ${targetAddress}`.trim(),
  renameWorksheetStep: (sheetName, newSheetName) => `Rename worksheet ${sheetName} to ${newSheetName}`.trim(),
  deleteWorksheetStep: (sheetName) => `Delete worksheet ${sheetName}`.trim(),
};

export function getUiStrings(locale: UiLocale): UiStrings {
  return locale === 'zh' ? zh : en;
}
```

Update `src/OfficeAgent.Frontend/src/components/ConfirmationCard.tsx` so it becomes locale-agnostic:

```tsx
type ConfirmationCardProps = {
  preview: ExcelCommandPreview;
  onConfirm: () => void;
  onCancel: () => void;
  ariaLabel: string;
  eyebrow: string;
  title: string;
  confirmLabel: string;
  cancelLabel: string;
};
```

```tsx
<article className="confirmation-card" aria-label={ariaLabel}>
  <div className="confirmation-card__eyebrow">{eyebrow}</div>
  <h2 className="confirmation-card__title">{title}</h2>
  <div className="confirmation-card__actions">
    <button type="button" className="ghost-button" onClick={onCancel}>{cancelLabel}</button>
    <button type="button" className="send-button" onClick={onConfirm}>{confirmLabel}</button>
  </div>
</article>
```

Update `src/OfficeAgent.Frontend/src/App.tsx` to fetch and use `HostContext`:

```tsx
import { getUiStrings, type UiLocale } from './i18n/uiStrings';

type BridgeStatusState =
  | { kind: 'connecting' }
  | { kind: 'connected'; host: string; version: string }
  | { kind: 'unavailable'; message: string };

const DEFAULT_SETTINGS: AppSettings = {
  apiKey: '',
  baseUrl: 'https://api.example.com',
  businessBaseUrl: '',
  model: 'gpt-5-mini',
  ssoUrl: '',
  ssoLoginSuccessPath: '',
  uiLanguageOverride: 'system',
};

export function App() {
  const [uiLocale, setUiLocale] = useState<UiLocale>('en');
  const strings = getUiStrings(uiLocale);
  const [bridgeStatus, setBridgeStatus] = useState<BridgeStatusState>({ kind: 'connecting' });
```

Call the host context before localizing the shell:

```tsx
useEffect(() => {
  let isActive = true;

  nativeBridge
    .getHostContext()
    .then((result) => {
      if (!isActive) {
        return;
      }

      setUiLocale(result.resolvedUiLocale);
    })
    .catch(() => {
      if (!isActive) {
        return;
      }

      setUiLocale('en');
    });
```

Store bridge status structurally so it can re-render when locale changes:

```tsx
nativeBridge
  .ping()
  .then((result) => {
    if (!isActive) {
      return;
    }

    setBridgeStatus({ kind: 'connected', host: result.host, version: result.version });
  })
  .catch((error: Error) => {
    if (!isActive) {
      return;
    }

    setBridgeStatus({ kind: 'unavailable', message: error.message });
  });
```

Localize untitled sessions and preserve backward compatibility with persisted English untitled threads:

```tsx
function isUntitledSession(title: string) {
  return title === 'New chat' || title === '新建会话';
}

function createUntitledSession(title: string): ChatSession {
  const now = new Date().toISOString();
  return {
    id: createMessageId(),
    title,
    createdAtUtc: now,
    updatedAtUtc: now,
    messages: [],
  };
}
```

Use `strings` for all fixed UI/system messages:

```tsx
const latestSession = allSessions[0];
let reusableSession: ChatSession | undefined;
if (latestSession && isUntitledSession(latestSession.title) && latestSession.messages.length === 0) {
  reusableSession = latestSession;
}
```

```tsx
appendThreadMessage(activeSession.id, {
  id: createMessageId(),
  role: 'system',
  content: activePendingConfirmation.kind === 'skill'
    ? strings.cancelPendingUpload
    : activePendingConfirmation.kind === 'agent'
      ? strings.cancelPendingPlan
      : strings.cancelPendingExcel,
});
```

```tsx
content: strings.requestFailed(error instanceof Error ? error.message : strings.excelCommandFailed),
```

Localize helper output:

```tsx
function createInitialThreadMessages(session: ChatSession | undefined, welcomeMessage: string): ThreadMessage[] {
  const persistedMessages = session?.messages ?? [];
  if (persistedMessages.length > 0) {
    return persistedMessages.map((message) => ({
      id: message.id,
      role: message.role === 'user' ? 'user' : 'assistant',
      content: message.content,
    }));
  }

  return [
    {
      id: 'welcome-message',
      role: 'assistant',
      content: welcomeMessage,
    },
  ];
}
```

```tsx
function createPlanPreview(result: AgentResult, strings: UiStrings): ExcelCommandPreview {
  const plan = result.planner?.plan;
  return {
    title: strings.executePlanTitle,
    summary: plan?.summary ?? result.message,
    details: plan?.steps.map((step) => formatPlanStep(step, strings)) ?? [],
  };
}
```

```tsx
function formatPlanStep(step: AgentPlan['steps'][number], strings: UiStrings) {
  switch (step.type) {
    case 'excel.addWorksheet':
      return strings.addWorksheetStep(String(step.args?.newSheetName ?? '').trim());
    case 'excel.writeRange':
      return strings.writeRangeStep(String(step.args?.targetAddress ?? '').trim());
    case 'excel.renameWorksheet':
      return strings.renameWorksheetStep(
        String(step.args?.sheetName ?? '').trim(),
        String(step.args?.newSheetName ?? '').trim(),
      );
    case 'excel.deleteWorksheet':
      return strings.deleteWorksheetStep(String(step.args?.sheetName ?? '').trim());
    case 'skill.upload_data':
      return strings.uploadSelectedData;
    default:
      return step.type;
  }
}
```

Wire localized props into `ConfirmationCard`:

```tsx
<ConfirmationCard
  preview={activePendingConfirmation.preview}
  onConfirm={() => void handlePendingConfirmationConfirm()}
  onCancel={handlePendingConfirmationCancel}
  ariaLabel={strings.confirmationAriaLabel}
  eyebrow={strings.confirmationEyebrow}
  title={strings.confirmationTitle}
  confirmLabel={strings.confirm}
  cancelLabel={strings.cancel}
/>
```

Finally, set the document language:

```tsx
useEffect(() => {
  document.documentElement.lang = uiLocale === 'zh' ? 'zh-CN' : 'en';
}, [uiLocale]);
```

Make browser-preview locale configurable in `src/OfficeAgent.Frontend/src/bridge/nativeBridge.ts`:

```ts
export class NativeBridge {
  constructor(
    webView: WebViewHostLike | undefined = getWebViewHost(),
    private readonly browserPreviewLocale: HostContext['resolvedUiLocale'] = 'en',
  ) {
    this.webView = webView;
    this.webView?.addEventListener('message', this.handleMessage);
  }
}
```

Use that locale in preview mode:

```ts
if (type === BRIDGE_TYPES.getHostContext) {
  return Promise.resolve({
    resolvedUiLocale: this.browserPreviewLocale,
    uiLanguageOverride: 'system',
  } as TResult);
}
```

Route the browser-preview helper text through the locale-aware string layer instead of hard-coded English:

```ts
import { getUiStrings } from '../i18n/uiStrings';

if (type === BRIDGE_TYPES.executeExcelCommand) {
  try {
    return Promise.resolve(
      createBrowserPreviewCommandResult(
        validateBrowserPreviewCommand(payload as ExcelCommand),
        this.browserPreviewLocale,
      ) as TResult,
    );
  } catch (error) {
    return Promise.reject(error);
  }
}
```

```ts
function createBrowserPreviewCommandResult(command: ExcelCommand, locale: 'zh' | 'en'): ExcelCommandResult {
  const strings = getUiStrings(locale);

  switch (command.commandType) {
    case 'excel.readSelectionTable':
      return {
        commandType: command.commandType,
        requiresConfirmation: false,
        status: 'completed',
        message: locale === 'zh' ? '已读取 Sheet1 A1:C4 的选区。' : 'Read selection from Sheet1 A1:C4.',
        table: {
          sheetName: 'Sheet1',
          address: 'A1:C4',
          headers: ['Name', 'Region', 'Amount'],
          rows: [
            ['Project A', 'CN', '42'],
            ['Project B', 'US', '36'],
          ],
        },
        selectionContext: BROWSER_PREVIEW_SELECTION_CONTEXT,
      };
    default:
      return createBrowserPreviewWriteResult(command, locale, strings);
  }
}
```

Use the explicit localized fallbacks in the three async command paths in `App.tsx`:

```tsx
content: strings.requestFailed(error instanceof Error ? error.message : strings.excelCommandFailed),
```

```tsx
content: strings.requestFailed(error instanceof Error ? error.message : strings.skillFailed),
```

```tsx
content: strings.requestFailed(error instanceof Error ? error.message : strings.agentFailed),
```

- [ ] **Step 4: Run the React tests and build to verify they pass**

Run:

```powershell
cd src/OfficeAgent.Frontend
npm run test -- src/App.test.tsx src/bridge/nativeBridge.test.ts
npm run build
```

Expected:

- Vitest PASS with `0` failed tests.
- `npm run build` completes without TypeScript or Vite errors.

- [ ] **Step 5: Commit**

```powershell
git add src/OfficeAgent.Frontend/index.html src/OfficeAgent.Frontend/src/i18n/uiStrings.ts src/OfficeAgent.Frontend/src/types/bridge.ts src/OfficeAgent.Frontend/src/bridge/nativeBridge.ts src/OfficeAgent.Frontend/src/components/ConfirmationCard.tsx src/OfficeAgent.Frontend/src/App.tsx src/OfficeAgent.Frontend/src/App.test.tsx src/OfficeAgent.Frontend/src/bridge/nativeBridge.test.ts
git commit -m "feat: localize task pane ui"
```

### Task 6: Update docs and run full bilingual verification

**Files:**
- Modify: `docs/modules/ribbon-sync-current-behavior.md`
- Create: `docs/modules/task-pane-current-behavior.md`
- Modify: `docs/module-index.md`
- Modify: `docs/vsto-manual-test-checklist.md`

- [ ] **Step 1: Update the module and manual-test docs**

Add a task-pane module snapshot in `docs/modules/task-pane-current-behavior.md`:

```markdown
# Task Pane Current Behavior

日期：2026-04-24

状态：已实现并可联调

## 1. 语言行为

- 默认根据 Excel UI 语言切换界面
- `zh-*` 显示中文 UI
- 其他语言显示英文 UI
- AI 自由回复仍尽量跟随用户输入语言

## 2. 覆盖范围

- 任务窗格固定 UI
- 欢迎语
- 任务窗格系统消息
- 执行计划标题与步骤文案
- 浏览器预览模式默认英文
```

Add a bilingual note to `docs/modules/ribbon-sync-current-behavior.md`:

```markdown
- Ribbon、登录弹窗、下载/上传确认框、初始化/错误提示会根据 Excel UI 语言自动切换中文或英文
- 仅 `zh-*` 显示中文；其余语言显示英文
```

Register the task-pane module in `docs/module-index.md`:

```markdown
| Task Pane | [docs/modules/task-pane-current-behavior.md](./modules/task-pane-current-behavior.md) | [docs/superpowers/specs/2026-04-24-office-agent-bilingual-ui-design.md](./superpowers/specs/2026-04-24-office-agent-bilingual-ui-design.md) | [src/OfficeAgent.Frontend/src/App.test.tsx](../src/OfficeAgent.Frontend/src/App.test.tsx) |
```

Extend `docs/vsto-manual-test-checklist.md` with two language passes:

```markdown
- 中文 Excel：
  - 打开 Ribbon，确认项目/下载/上传/账号分组为中文
  - 打开任务窗格，确认欢迎语、设置、会话抽屉、确认卡片为中文
- 英文 Excel：
  - 打开 Ribbon，确认 host UI 为英文
  - 打开任务窗格，确认固定 UI 与 host error 文案为英文
  - 输入中文自由提问，确认 AI 回复仍可按中文输出
```

- [ ] **Step 2: Run the full automated verification**

Run:

```powershell
dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj
cd src/OfficeAgent.Frontend
npm run test
npm run build
```

Expected: all commands PASS with `0` failed tests and no TypeScript build errors.

- [ ] **Step 3: Rebuild the add-in and perform the bilingual smoke test**

Run:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File eng/Dev-RefreshExcelAddIn.ps1 -CloseExcel
```

Expected: the script completes successfully after rebuilding the frontend, rebuilding the add-in, and refreshing the local Excel registration. Then manually validate one Chinese Excel session and one English Excel session using the updated checklist.

- [ ] **Step 4: Commit the docs and verification artifacts**

```powershell
git add docs/modules/ribbon-sync-current-behavior.md docs/modules/task-pane-current-behavior.md docs/module-index.md docs/vsto-manual-test-checklist.md
git commit -m "docs: document bilingual ui behavior"
```
