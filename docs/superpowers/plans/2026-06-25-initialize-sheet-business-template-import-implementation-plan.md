# Initialize Sheet Business Template Import Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Let `初始化当前表` optionally create the current worksheet from a business-system exported `.xlsx` template while preserving the existing config-only initialization path.

**Architecture:** Keep the business contract in Core as an optional connector extension, keep HTTP details in Infrastructure, and keep Excel COM copy plus WinForms interaction in ExcelAddIn. Metadata is prepared before worksheet mutation and written only after the `Business Data` sheet is copied, so download/pre-copy failures leave the workbook unchanged.

**Tech Stack:** C# .NET Framework 4.8, Excel VSTO COM interop, WinForms dialogs, xUnit tests, Node/Express mock server with the `xlsx` package for local binary workbook export.

---

## Confirmed Product Contract

- UI text may continue to say `模板`.
- Code and docs must distinguish `Business Export Template` from the existing local sync configuration template.
- Business export workbook source sheet name is exactly `Business Data`.
- Blank work sheet defaults to template import and preserves the current sheet name.
- Nonblank work sheet defaults to config-only; template import is allowed only after an explicit mode switch and visible overwrite warning.
- Blank detection is content-only. Formatting, row height, column width, frozen panes, and UsedRange expansion do not make the sheet nonblank.
- Project selection stays in the Ribbon project dropdown. The initialization dialog does not select a project.
- Template list comes from the business system and contains only `templateId` and `templateName`.
- Template import does not record `templateId` or `templateName` in `xISDP_Setting`.
- Template download returns binary `.xlsx`; no URL flow is supported.
- Cancellation is guaranteed only during HTTP download. COM copy cancellation is not supported in this version.
- `xISDP_Setting` and `xISDP_Log` are protected from initialization and template import.
- Full download is outside this implementation. Users can select sheet data and click the existing Download button.

## File Structure

Create:

- `src/OfficeAgent.Core/Models/BusinessExportTemplateOption.cs`  
  Stable business template list item: `TemplateId`, `TemplateName`.
- `src/OfficeAgent.Core/Models/BusinessExportWorkbook.cs`  
  Downloaded workbook binary plus file name and content type.
- `src/OfficeAgent.Core/Models/SheetInitializationPlan.cs`  
  Prepared binding, mapping definition, and field mapping rows before writing metadata.
- `src/OfficeAgent.Core/Services/IBusinessExportTemplateConnector.cs`  
  Optional connector extension. It is not added to `ISystemConnector`.
- `src/OfficeAgent.ExcelAddIn/Excel/IBusinessWorkbookImporter.cs`  
  Testable boundary for blank detection, write protection preflight, workbook copy, and focus restoration.
- `src/OfficeAgent.ExcelAddIn/Excel/ExcelBusinessWorkbookImporter.cs`  
  Excel COM implementation that opens the downloaded `.xlsx`, copies `Business Data` into the current worksheet, preserves the current worksheet name, and deletes the temp file.
- `src/OfficeAgent.ExcelAddIn/Dialogs/InitializeSheetDialogModels.cs`  
  Dialog request/result/load result models and import mode enum.
- `src/OfficeAgent.ExcelAddIn/Dialogs/InitializeSheetDialog.cs`  
  WinForms initialization dialog with loading, disabled template states, config-only mode, template mode, and overwrite warning.
- `src/OfficeAgent.ExcelAddIn/Dialogs/InitializeSheetImportProgressDialog.cs`  
  Modal progress dialog with cancel enabled only during download.
- `tests/OfficeAgent.ExcelAddIn.Tests/InitializeSheetDialogTests.cs`  
  Dialog model/default-mode tests without showing a real modal dialog.
- `tests/OfficeAgent.ExcelAddIn.Tests/ExcelBusinessWorkbookImporterConfigurationTests.cs`  
  Source-level guard tests for COM copy contract, temp-file deletion, and managed source sheet name.

Modify:

- `src/OfficeAgent.Core/Sync/WorksheetSyncService.cs`  
  Add deferred initialization plan creation, metadata save method, optional template list/export methods, and analytics.
- `src/OfficeAgent.Infrastructure/Http/CurrentBusinessSystemConnector.cs`  
  Implement `IBusinessExportTemplateConnector` with `POST /templates` and `POST /export`.
- `src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs`  
  Add blank detection, template list loading, and template import orchestration.
- `src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs`  
  Route `ExecuteInitializeCurrentSheet` through the new dialog and progress flow.
- `src/OfficeAgent.ExcelAddIn/Dialogs/OperationResultDialog.cs`  
  Extend `IRibbonSyncDialogService` with initialization dialog and progress methods.
- `src/OfficeAgent.ExcelAddIn/Localization/HostLocalizedStrings.cs`  
  Add all new dialog, progress, warning, success, and failure strings.
- `src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj`  
  Include new Excel and dialog files.
- `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs`  
  Construct `ExcelBusinessWorkbookImporter` and pass it to `WorksheetSyncExecutionService`.
- `tests/OfficeAgent.Core.Tests/WorksheetSyncServiceTests.cs`  
  Add deferred metadata and optional connector tests.
- `tests/OfficeAgent.Infrastructure.Tests/CurrentBusinessSystemConnectorTests.cs`  
  Add template list, binary export, auth, and cancellation tests.
- `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs`  
  Add template import orchestration tests using fake importer and fake connector.
- `tests/OfficeAgent.ExcelAddIn.Tests/RibbonSyncControllerTests.cs`  
  Add initialization dialog, managed sheet, success copy, cancel, and analytics tests.
- `tests/OfficeAgent.ExcelAddIn.Tests/HostLocalizedStringsTests.cs`  
  Add localization coverage for new strings.
- `tests/mock-server/package.json` and `tests/mock-server/package-lock.json`  
  Add `xlsx` dependency.
- `tests/mock-server/server.js`  
  Add `POST /templates` and `POST /export`.
- `tests/mock-server/README.md`  
  Document the new mock endpoints.
- `docs/modules/ribbon-sync-current-behavior.md`  
  Update current user-visible behavior after implementation.
- `docs/ribbon-sync-real-system-integration-guide.md`  
  Document optional business template export connector.
- `docs/vsto-manual-test-checklist.md`  
  Add manual coverage for blank/nonblank initialization and cancellation.

Chosen business API endpoints:

```text
POST /templates
Request:  { "projectId": "performance" }
Response: [
  { "templateId": "standard", "templateName": "标准作业表" }
]

POST /export
Request:  { "projectId": "performance", "templateId": "standard" }
Response: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet bytes
```

---

### Task 1: Core Optional Template Contract And Deferred Initialization

**Files:**
- Create: `src/OfficeAgent.Core/Models/BusinessExportTemplateOption.cs`
- Create: `src/OfficeAgent.Core/Models/BusinessExportWorkbook.cs`
- Create: `src/OfficeAgent.Core/Models/SheetInitializationPlan.cs`
- Create: `src/OfficeAgent.Core/Services/IBusinessExportTemplateConnector.cs`
- Modify: `src/OfficeAgent.Core/Sync/WorksheetSyncService.cs`
- Test: `tests/OfficeAgent.Core.Tests/WorksheetSyncServiceTests.cs`

- [ ] **Step 1: Add failing Core tests for deferred initialization writes**

Add these tests to `tests/OfficeAgent.Core.Tests/WorksheetSyncServiceTests.cs`:

```csharp
[Fact]
public void PrepareSheetInitializationBuildsPlanWithoutWritingMetadata()
{
    var connector = new FakeSystemConnector();
    var metadataStore = new FakeWorksheetMetadataStore();
    var service = CreateService(connector, metadataStore);
    var project = new ProjectOption
    {
        SystemKey = connector.SystemKey,
        ProjectId = "performance",
        DisplayName = "绩效项目",
    };

    var plan = service.PrepareSheetInitialization("Sheet1", project);

    Assert.Equal("Sheet1", plan.Binding.SheetName);
    Assert.Equal("performance", plan.Binding.ProjectId);
    Assert.Same(connector.FieldMappingDefinition, plan.FieldMappingDefinition);
    Assert.Equal(connector.FieldMappingSeedRows.Count, plan.FieldMappings.Count);
    Assert.Null(metadataStore.LastSavedBinding);
    Assert.Empty(metadataStore.LastSavedFieldMappings);
}

[Fact]
public void SaveSheetInitializationWritesPreparedBindingAndMappings()
{
    var connector = new FakeSystemConnector();
    var metadataStore = new FakeWorksheetMetadataStore();
    var service = CreateService(connector, metadataStore);
    var project = new ProjectOption
    {
        SystemKey = connector.SystemKey,
        ProjectId = "performance",
        DisplayName = "绩效项目",
    };
    var plan = service.PrepareSheetInitialization("Sheet1", project);

    service.SaveSheetInitialization(plan);

    Assert.Equal("Sheet1", metadataStore.LastSavedBinding.SheetName);
    Assert.Same(connector.FieldMappingDefinition, metadataStore.LastSavedFieldMappingDefinition);
    Assert.Equal(
        connector.FieldMappingSeedRows.Select(row => row.Values["ApiFieldKey"]),
        metadataStore.LastSavedFieldMappings.Select(row => row.Values["ApiFieldKey"]));
}

[Fact]
public void InitializeSheetUsesDeferredPlanAndPreservesExistingBehavior()
{
    var connector = new FakeSystemConnector();
    var metadataStore = new FakeWorksheetMetadataStore();
    var service = CreateService(connector, metadataStore);

    service.InitializeSheet("Sheet1", new ProjectOption
    {
        SystemKey = connector.SystemKey,
        ProjectId = "performance",
        DisplayName = "绩效项目",
    });

    Assert.Equal("Sheet1", metadataStore.LastSavedBinding.SheetName);
    Assert.Equal("performance", connector.LastFieldMappingDefinitionProjectId);
    Assert.NotEmpty(metadataStore.LastSavedFieldMappings);
}
```

- [ ] **Step 2: Run Core tests and verify the expected failures**

Run:

```powershell
dotnet test tests/OfficeAgent.Core.Tests/OfficeAgent.Core.Tests.csproj --filter "FullyQualifiedName~PrepareSheetInitializationBuildsPlanWithoutWritingMetadata|FullyQualifiedName~SaveSheetInitializationWritesPreparedBindingAndMappings|FullyQualifiedName~InitializeSheetUsesDeferredPlanAndPreservesExistingBehavior"
```

Expected: the first two tests fail because `PrepareSheetInitialization`, `SaveSheetInitialization`, and `SheetInitializationPlan` do not exist.

- [ ] **Step 3: Add Core business export models**

Create `src/OfficeAgent.Core/Models/BusinessExportTemplateOption.cs`:

```csharp
namespace OfficeAgent.Core.Models
{
    public sealed class BusinessExportTemplateOption
    {
        public string TemplateId { get; set; } = string.Empty;

        public string TemplateName { get; set; } = string.Empty;
    }
}
```

Create `src/OfficeAgent.Core/Models/BusinessExportWorkbook.cs`:

```csharp
using System;

namespace OfficeAgent.Core.Models
{
    public sealed class BusinessExportWorkbook
    {
        public string FileName { get; set; } = "business-export.xlsx";

        public string ContentType { get; set; } = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        public byte[] Content { get; set; } = Array.Empty<byte>();
    }
}
```

Create `src/OfficeAgent.Core/Models/SheetInitializationPlan.cs`:

```csharp
using System;
using System.Collections.Generic;

namespace OfficeAgent.Core.Models
{
    public sealed class SheetInitializationPlan
    {
        public SheetBinding Binding { get; set; }

        public FieldMappingTableDefinition FieldMappingDefinition { get; set; }

        public IReadOnlyList<SheetFieldMappingRow> FieldMappings { get; set; } = Array.Empty<SheetFieldMappingRow>();
    }
}
```

Create `src/OfficeAgent.Core/Services/IBusinessExportTemplateConnector.cs`:

```csharp
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Services
{
    public interface IBusinessExportTemplateConnector
    {
        IReadOnlyList<BusinessExportTemplateOption> GetBusinessExportTemplates(string projectId);

        Task<BusinessExportWorkbook> ExportBusinessWorkbookAsync(
            string projectId,
            string templateId,
            CancellationToken cancellationToken);
    }
}
```

- [ ] **Step 4: Split `WorksheetSyncService.InitializeSheet` into prepare and save**

In `src/OfficeAgent.Core/Sync/WorksheetSyncService.cs`, replace the body of `InitializeSheet` with:

```csharp
public void InitializeSheet(string sheetName, ProjectOption project)
{
    var stopwatch = Stopwatch.StartNew();
    var properties = BuildConnectorProperties(project?.SystemKey, project?.ProjectId);
    try
    {
        var plan = PrepareSheetInitialization(sheetName, project);
        SaveSheetInitialization(plan);
        properties["fieldMappingColumnCount"] = plan.FieldMappingDefinition?.Columns?.Length ?? 0;
        properties["fieldMappingRowCount"] = plan.FieldMappings?.Count ?? 0;
        TrackConnectorEvent("connector.initialize_sheet.completed", properties, stopwatch);
    }
    catch (Exception ex)
    {
        TrackConnectorEvent("connector.initialize_sheet.failed", properties, stopwatch, ToAnalyticsError(ex));
        throw;
    }
}
```

Add these public methods below `InitializeSheet`:

```csharp
public SheetInitializationPlan PrepareSheetInitialization(string sheetName, ProjectOption project)
{
    if (string.IsNullOrWhiteSpace(sheetName))
    {
        throw new ArgumentException("Sheet name is required.", nameof(sheetName));
    }

    if (project == null)
    {
        throw new ArgumentNullException(nameof(project));
    }

    var connector = GetRequiredConnector(project.SystemKey);
    var bindingSeed = connector.CreateBindingSeed(sheetName, project);
    var binding = MergeExistingLayout(bindingSeed);
    var definition = connector.GetFieldMappingDefinition(project.ProjectId);
    var seedRows = connector.BuildFieldMappingSeed(sheetName, project.ProjectId);

    return new SheetInitializationPlan
    {
        Binding = binding,
        FieldMappingDefinition = definition,
        FieldMappings = seedRows ?? Array.Empty<SheetFieldMappingRow>(),
    };
}

public void SaveSheetInitialization(SheetInitializationPlan plan)
{
    if (plan == null)
    {
        throw new ArgumentNullException(nameof(plan));
    }

    if (plan.Binding == null)
    {
        throw new InvalidOperationException("Sheet initialization binding is required.");
    }

    metadataStore.SaveBinding(plan.Binding);
    metadataStore.SaveFieldMappings(
        plan.Binding.SheetName,
        plan.FieldMappingDefinition,
        plan.FieldMappings ?? Array.Empty<SheetFieldMappingRow>());
}
```

- [ ] **Step 5: Add optional business export methods to `WorksheetSyncService`**

Add `using System.Threading;` and `using System.Threading.Tasks;` to `WorksheetSyncService.cs`.

Add these public methods below `SaveSheetInitialization`:

```csharp
public bool SupportsBusinessExportTemplates(string systemKey)
{
    return GetRequiredConnector(systemKey) is IBusinessExportTemplateConnector;
}

public IReadOnlyList<BusinessExportTemplateOption> GetBusinessExportTemplates(string systemKey, string projectId)
{
    var connector = GetRequiredConnector(systemKey);
    if (!(connector is IBusinessExportTemplateConnector templateConnector))
    {
        return Array.Empty<BusinessExportTemplateOption>();
    }

    return (templateConnector.GetBusinessExportTemplates(projectId) ?? Array.Empty<BusinessExportTemplateOption>())
        .Where(template => template != null && !string.IsNullOrWhiteSpace(template.TemplateId))
        .Select(template => new BusinessExportTemplateOption
        {
            TemplateId = template.TemplateId ?? string.Empty,
            TemplateName = string.IsNullOrWhiteSpace(template.TemplateName)
                ? template.TemplateId ?? string.Empty
                : template.TemplateName,
        })
        .ToArray();
}

public Task<BusinessExportWorkbook> ExportBusinessWorkbookAsync(
    string systemKey,
    string projectId,
    string templateId,
    CancellationToken cancellationToken)
{
    if (string.IsNullOrWhiteSpace(templateId))
    {
        throw new ArgumentException("Template id is required.", nameof(templateId));
    }

    var connector = GetRequiredConnector(systemKey);
    if (!(connector is IBusinessExportTemplateConnector templateConnector))
    {
        throw new InvalidOperationException("The current business system does not support template export.");
    }

    return templateConnector.ExportBusinessWorkbookAsync(projectId, templateId, cancellationToken);
}
```

- [ ] **Step 6: Add Core tests for optional connector support**

In `tests/OfficeAgent.Core.Tests/WorksheetSyncServiceTests.cs`, add:

```csharp
[Fact]
public void SupportsBusinessExportTemplatesReturnsFalseWhenConnectorDoesNotImplementExtension()
{
    var service = CreateService(new FakeSystemConnector(), new FakeWorksheetMetadataStore());

    Assert.False(service.SupportsBusinessExportTemplates("current-business-system"));
    Assert.Empty(service.GetBusinessExportTemplates("current-business-system", "performance"));
}

[Fact]
public async System.Threading.Tasks.Task ExportBusinessWorkbookAsyncUsesOptionalConnector()
{
    var connector = new FakeBusinessTemplateConnector();
    var service = CreateService(connector, new FakeWorksheetMetadataStore());

    var workbook = await service.ExportBusinessWorkbookAsync(
        connector.SystemKey,
        "performance",
        "standard",
        System.Threading.CancellationToken.None);

    Assert.Equal("performance", connector.LastExportProjectId);
    Assert.Equal("standard", connector.LastExportTemplateId);
    Assert.Equal(new byte[] { 0x50, 0x4B }, workbook.Content);
}
```

In the same file, change the existing fake connector declaration from:

```csharp
private sealed class FakeSystemConnector : ISystemConnector, IUploadChangeFilter
```

to:

```csharp
private class FakeSystemConnector : ISystemConnector, IUploadChangeFilter
```

Add this fake connector class inside the test class:

```csharp
private sealed class FakeBusinessTemplateConnector : FakeSystemConnector, IBusinessExportTemplateConnector
{
    public string LastExportProjectId { get; private set; }

    public string LastExportTemplateId { get; private set; }

    public IReadOnlyList<BusinessExportTemplateOption> GetBusinessExportTemplates(string projectId)
    {
        return new[]
        {
            new BusinessExportTemplateOption
            {
                TemplateId = "standard",
                TemplateName = "标准作业表",
            },
        };
    }

    public System.Threading.Tasks.Task<BusinessExportWorkbook> ExportBusinessWorkbookAsync(
        string projectId,
        string templateId,
        System.Threading.CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        LastExportProjectId = projectId;
        LastExportTemplateId = templateId;
        return System.Threading.Tasks.Task.FromResult(new BusinessExportWorkbook
        {
            FileName = "standard.xlsx",
            ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            Content = new byte[] { 0x50, 0x4B },
        });
    }
}
```

- [ ] **Step 7: Run Core tests**

Run:

```powershell
dotnet test tests/OfficeAgent.Core.Tests/OfficeAgent.Core.Tests.csproj
```

Expected: PASS.

- [ ] **Step 8: Commit Core contract**

Run:

```powershell
git add src/OfficeAgent.Core/Models/BusinessExportTemplateOption.cs src/OfficeAgent.Core/Models/BusinessExportWorkbook.cs src/OfficeAgent.Core/Models/SheetInitializationPlan.cs src/OfficeAgent.Core/Services/IBusinessExportTemplateConnector.cs src/OfficeAgent.Core/Sync/WorksheetSyncService.cs tests/OfficeAgent.Core.Tests/WorksheetSyncServiceTests.cs
git commit -m "feat: add business export template contract"
```

---

### Task 2: Current Business Connector Template List And Binary Export

**Files:**
- Modify: `src/OfficeAgent.Infrastructure/Http/CurrentBusinessSystemConnector.cs`
- Test: `tests/OfficeAgent.Infrastructure.Tests/CurrentBusinessSystemConnectorTests.cs`

- [ ] **Step 1: Add failing connector tests**

Add these tests to `tests/OfficeAgent.Infrastructure.Tests/CurrentBusinessSystemConnectorTests.cs`:

```csharp
using OfficeAgent.Core.Services;

[Fact]
public void GetBusinessExportTemplatesCallsTemplatesEndpoint()
{
    var handler = new TemplateExportHandler();
    var connector = CurrentBusinessSystemConnector.ForTests("https://api.internal.example", handler);

    var templates = ((IBusinessExportTemplateConnector)connector).GetBusinessExportTemplates("performance");

    Assert.Equal("/templates", handler.Requests.Single().Path);
    Assert.Contains("\"projectId\":\"performance\"", handler.Requests.Single().Body);
    var template = Assert.Single(templates);
    Assert.Equal("standard", template.TemplateId);
    Assert.Equal("标准作业表", template.TemplateName);
}

[Fact]
public async Task ExportBusinessWorkbookAsyncDownloadsXlsxBytes()
{
    var handler = new TemplateExportHandler();
    var connector = CurrentBusinessSystemConnector.ForTests("https://api.internal.example", handler);

    var workbook = await ((IBusinessExportTemplateConnector)connector).ExportBusinessWorkbookAsync(
        "performance",
        "standard",
        CancellationToken.None);

    Assert.Equal("/export", handler.Requests.Last().Path);
    Assert.Contains("\"projectId\":\"performance\"", handler.Requests.Last().Body);
    Assert.Contains("\"templateId\":\"standard\"", handler.Requests.Last().Body);
    Assert.Equal(new byte[] { 0x50, 0x4B, 0x03, 0x04 }, workbook.Content);
    Assert.Equal("business-export.xlsx", workbook.FileName);
}

[Theory]
[InlineData(HttpStatusCode.Unauthorized)]
[InlineData(HttpStatusCode.Forbidden)]
public async Task ExportBusinessWorkbookAsyncPromptsLoginForAuthenticationFailures(HttpStatusCode statusCode)
{
    var connector = CurrentBusinessSystemConnector.ForTests(
        "https://api.internal.example",
        new TemplateExportHandler(_ => new HttpResponseMessage(statusCode)
        {
            Content = new StringContent("{\"code\":\"unauthorized\"}", Encoding.UTF8, "application/json"),
        }));

    await Assert.ThrowsAsync<AuthenticationRequiredException>(() =>
        ((IBusinessExportTemplateConnector)connector).ExportBusinessWorkbookAsync(
            "performance",
            "standard",
            CancellationToken.None));
}

[Fact]
public async Task ExportBusinessWorkbookAsyncPropagatesCancellation()
{
    var connector = CurrentBusinessSystemConnector.ForTests(
        "https://api.internal.example",
        new CanceledTemplateExportHandler());
    using (var cancellationTokenSource = new CancellationTokenSource())
    {
        cancellationTokenSource.Cancel();

        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            ((IBusinessExportTemplateConnector)connector).ExportBusinessWorkbookAsync(
                "performance",
                "standard",
                cancellationTokenSource.Token));
    }
}
```

Add these handlers to the test file:

```csharp
private sealed class TemplateExportHandler : HttpMessageHandler
{
    private readonly Func<HttpRequestMessage, HttpResponseMessage> createResponse;

    public TemplateExportHandler(Func<HttpRequestMessage, HttpResponseMessage> createResponse = null)
    {
        this.createResponse = createResponse;
    }

    public List<(string Path, string Body)> Requests { get; } = new List<(string Path, string Body)>();

    protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
    {
        var body = request.Content?.ReadAsStringAsync().GetAwaiter().GetResult() ?? string.Empty;
        Requests.Add((request.RequestUri?.AbsolutePath ?? string.Empty, body));
        if (createResponse != null)
        {
            return Task.FromResult(createResponse(request));
        }

        if (request.RequestUri?.AbsolutePath == "/templates")
        {
            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(
                    "[{\"templateId\":\"standard\",\"templateName\":\"标准作业表\"}]",
                    Encoding.UTF8,
                    "application/json"),
            });
        }

        return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
        {
            Content = new ByteArrayContent(new byte[] { 0x50, 0x4B, 0x03, 0x04 }),
        });
    }
}

private sealed class CanceledTemplateExportHandler : HttpMessageHandler
{
    protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        throw new OperationCanceledException(cancellationToken);
    }
}
```

- [ ] **Step 2: Run connector tests and verify expected failures**

Run:

```powershell
dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj --filter "FullyQualifiedName~GetBusinessExportTemplatesCallsTemplatesEndpoint|FullyQualifiedName~ExportBusinessWorkbookAsync"
```

Expected: FAIL because `CurrentBusinessSystemConnector` does not implement `IBusinessExportTemplateConnector`.

- [ ] **Step 3: Implement template list and binary export**

In `src/OfficeAgent.Infrastructure/Http/CurrentBusinessSystemConnector.cs`, change the class declaration:

```csharp
public sealed class CurrentBusinessSystemConnector : ISystemConnector, IBusinessExportTemplateConnector
```

Add constants near the existing defaults:

```csharp
private const string TemplatesPath = "/templates";
private const string ExportPath = "/export";
private const string XlsxContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
```

Add these public methods below `GetProjects`:

```csharp
public IReadOnlyList<BusinessExportTemplateOption> GetBusinessExportTemplates(string projectId)
{
    var stopwatch = Stopwatch.StartNew();
    var properties = BuildBusinessProperties(projectId);
    try
    {
        EnsureProjectId(projectId);
        var templates = Post<List<BusinessExportTemplateOption>>(TemplatesPath, new { projectId })
            ?? new List<BusinessExportTemplateOption>();
        var normalized = templates
            .Where(template => template != null && !string.IsNullOrWhiteSpace(template.TemplateId))
            .Select(template => new BusinessExportTemplateOption
            {
                TemplateId = template.TemplateId ?? string.Empty,
                TemplateName = string.IsNullOrWhiteSpace(template.TemplateName)
                    ? template.TemplateId ?? string.Empty
                    : template.TemplateName,
            })
            .ToArray();

        properties["templateCount"] = normalized.Length;
        TrackBusinessEvent("business.current.templates.completed", properties, TemplatesPath, "templates", stopwatch);
        return normalized;
    }
    catch (Exception ex)
    {
        TrackBusinessEvent("business.current.templates.failed", properties, TemplatesPath, "templates", stopwatch, ToAnalyticsError(ex));
        throw;
    }
}

public async Task<BusinessExportWorkbook> ExportBusinessWorkbookAsync(
    string projectId,
    string templateId,
    CancellationToken cancellationToken)
{
    var stopwatch = Stopwatch.StartNew();
    var properties = BuildBusinessProperties(projectId);
    properties["templateId"] = templateId ?? string.Empty;
    try
    {
        EnsureProjectId(projectId);
        if (string.IsNullOrWhiteSpace(templateId))
        {
            throw new InvalidOperationException("Template id is required for current business system.");
        }

        using (var response = await SendAsync(HttpMethod.Post, ExportPath, new { projectId, templateId }, cancellationToken).ConfigureAwait(false))
        {
            var content = response.Content == null
                ? Array.Empty<byte>()
                : await response.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
            EnsureSuccessStatusCode(response, string.Empty);
            if (content.Length == 0)
            {
                throw new InvalidOperationException("Business export returned an empty workbook.");
            }

            properties["byteCount"] = content.Length;
            TrackBusinessEvent("business.current.export_workbook.completed", properties, ExportPath, "export_workbook", stopwatch);
            return new BusinessExportWorkbook
            {
                FileName = "business-export.xlsx",
                ContentType = response.Content?.Headers?.ContentType?.MediaType ?? XlsxContentType,
                Content = content,
            };
        }
    }
    catch (OperationCanceledException)
    {
        TrackBusinessEvent("business.current.export_workbook.canceled", properties, ExportPath, "export_workbook", stopwatch);
        throw;
    }
    catch (Exception ex)
    {
        TrackBusinessEvent("business.current.export_workbook.failed", properties, ExportPath, "export_workbook", stopwatch, ToAnalyticsError(ex));
        throw;
    }
}
```

Add `using System.Threading;` and `using System.Threading.Tasks;`.

- [ ] **Step 4: Add cancellable HTTP helper**

Replace `Send` with a wrapper around a new async method:

```csharp
private HttpResponseMessage Send(HttpMethod method, string path, object payload)
{
    return SendAsync(method, path, payload, CancellationToken.None).GetAwaiter().GetResult();
}

private async Task<HttpResponseMessage> SendAsync(
    HttpMethod method,
    string path,
    object payload,
    CancellationToken cancellationToken)
{
    var baseUri = ResolveBaseUri();
    using (var request = new HttpRequestMessage(method, new Uri(baseUri, path)))
    {
        var projectId = ExtractProjectId(payload);
        if (payload != null)
        {
            request.Content = new StringContent(JsonConvert.SerializeObject(payload), Encoding.UTF8, "application/json");
        }

        OfficeAgentLog.Info(
            "business_api",
            "request.begin",
            "Business API request started.",
            BuildRequestDetails(method, path, projectId));
        try
        {
            var response = await httpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false);
            OfficeAgentLog.Info(
                "business_api",
                "request.completed",
                $"Business API request completed with {(int)response.StatusCode} {response.ReasonPhrase}.",
                BuildRequestDetails(method, path, projectId));
            return response;
        }
        catch (OperationCanceledException ex)
        {
    OfficeAgentLog.Warn(
        "business_api",
        "request.canceled",
        "Business API request was canceled.",
        BuildRequestDetails(method, path, projectId));
            throw;
        }
        catch (HttpRequestException ex)
        {
            OfficeAgentLog.Error(
                "business_api",
                "request.exception",
                "Business API request failed with an HTTP transport error.",
                ex,
                BuildRequestDetails(method, path, projectId));
            throw;
        }
    }
}
```

Keep `Post<T>`, `Get<T>`, `PostBatchSave`, and existing sync callers unchanged. They continue to call `Send`.

- [ ] **Step 5: Run Infrastructure tests**

Run:

```powershell
dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj
```

Expected: PASS.

- [ ] **Step 6: Commit connector export**

Run:

```powershell
git add src/OfficeAgent.Infrastructure/Http/CurrentBusinessSystemConnector.cs tests/OfficeAgent.Infrastructure.Tests/CurrentBusinessSystemConnectorTests.cs
git commit -m "feat: download business template workbooks"
```

---

### Task 3: Excel Workbook Import Boundary And COM Implementation

**Files:**
- Create: `src/OfficeAgent.ExcelAddIn/Excel/IBusinessWorkbookImporter.cs`
- Create: `src/OfficeAgent.ExcelAddIn/Excel/ExcelBusinessWorkbookImporter.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj`
- Test: `tests/OfficeAgent.ExcelAddIn.Tests/ExcelBusinessWorkbookImporterConfigurationTests.cs`

- [ ] **Step 1: Add source-level tests for importer contract**

Create `tests/OfficeAgent.ExcelAddIn.Tests/ExcelBusinessWorkbookImporterConfigurationTests.cs`:

```csharp
using System;
using System.IO;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class ExcelBusinessWorkbookImporterConfigurationTests
    {
        [Fact]
        public void ImporterUsesBusinessDataSheetNameAndPreservesTargetSheetName()
        {
            var text = ReadSource("src", "OfficeAgent.ExcelAddIn", "Excel", "ExcelBusinessWorkbookImporter.cs");

            Assert.Contains("Business Data", text, StringComparison.Ordinal);
            Assert.Contains("var originalTargetSheetName = targetWorksheet.Name;", text, StringComparison.Ordinal);
            Assert.Contains("targetWorksheet.Name = originalTargetSheetName;", text, StringComparison.Ordinal);
        }

        [Fact]
        public void ImporterDeletesTemporaryWorkbookInFinallyBlock()
        {
            var text = ReadSource("src", "OfficeAgent.ExcelAddIn", "Excel", "ExcelBusinessWorkbookImporter.cs");

            Assert.Contains("finally", text, StringComparison.Ordinal);
            Assert.Contains("File.Delete(tempPath);", text, StringComparison.Ordinal);
        }

        [Fact]
        public void ImporterDetectsContentWithConstantsAndFormulas()
        {
            var text = ReadSource("src", "OfficeAgent.ExcelAddIn", "Excel", "ExcelBusinessWorkbookImporter.cs");

            Assert.Contains("xlCellTypeConstants", text, StringComparison.Ordinal);
            Assert.Contains("xlCellTypeFormulas", text, StringComparison.Ordinal);
        }

        private static string ReadSource(params string[] segments)
        {
            return File.ReadAllText(ResolveRepositoryPath(segments));
        }

        private static string ResolveRepositoryPath(params string[] segments)
        {
            return Path.GetFullPath(Path.Combine(new[]
            {
                AppContext.BaseDirectory,
                "..",
                "..",
                "..",
                "..",
                "..",
            }.Concat(segments).ToArray()));
        }
    }
}
```

Add `using System.Linq;` at the top of that file.

- [ ] **Step 2: Run importer tests and verify expected failures**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~ExcelBusinessWorkbookImporterConfigurationTests"
```

Expected: FAIL because `ExcelBusinessWorkbookImporter.cs` does not exist.

- [ ] **Step 3: Add importer interface**

Create `src/OfficeAgent.ExcelAddIn/Excel/IBusinessWorkbookImporter.cs`:

```csharp
namespace OfficeAgent.ExcelAddIn.Excel
{
    internal interface IBusinessWorkbookImporter
    {
        bool IsWorkSheetContentBlank(string sheetName);

        void EnsureCanWriteToWorkSheet(string sheetName);

        void ImportBusinessDataSheet(byte[] workbookBytes, string targetSheetName);

        void ActivateWorkSheetAtA1(string sheetName);
    }
}
```

- [ ] **Step 4: Add COM importer implementation**

Create `src/OfficeAgent.ExcelAddIn/Excel/ExcelBusinessWorkbookImporter.cs`:

```csharp
using System;
using System.IO;
using System.Runtime.InteropServices;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal sealed class ExcelBusinessWorkbookImporter : IBusinessWorkbookImporter
    {
        private const string BusinessDataSheetName = "Business Data";
        private readonly ExcelInterop.Application application;

        public ExcelBusinessWorkbookImporter(ExcelInterop.Application application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public bool IsWorkSheetContentBlank(string sheetName)
        {
            var worksheet = GetWorksheet(sheetName);
            return !HasSpecialCells(worksheet, ExcelInterop.XlCellType.xlCellTypeConstants) &&
                   !HasSpecialCells(worksheet, ExcelInterop.XlCellType.xlCellTypeFormulas);
        }

        public void EnsureCanWriteToWorkSheet(string sheetName)
        {
            var workbook = GetWorkbook();
            var worksheet = GetWorksheet(sheetName);
            if (workbook.ProtectStructure)
            {
                throw new InvalidOperationException("The workbook structure is protected and cannot receive template content.");
            }

            if (worksheet.ProtectContents || worksheet.ProtectDrawingObjects || worksheet.ProtectScenarios)
            {
                throw new InvalidOperationException("The current worksheet is protected and cannot receive template content.");
            }
        }

        public void ImportBusinessDataSheet(byte[] workbookBytes, string targetSheetName)
        {
            if (workbookBytes == null || workbookBytes.Length == 0)
            {
                throw new InvalidOperationException("Business export workbook is empty.");
            }

            if (workbookBytes.Length < 2 || workbookBytes[0] != 0x50 || workbookBytes[1] != 0x4B)
            {
                throw new InvalidOperationException("Business export workbook is not a valid .xlsx file.");
            }

            EnsureCanWriteToWorkSheet(targetSheetName);

            var tempPath = Path.Combine(
                Path.GetTempPath(),
                "OfficeAgent-BusinessExport-" + Guid.NewGuid().ToString("N") + ".xlsx");
            ExcelInterop.Workbook sourceWorkbook = null;
            try
            {
                File.WriteAllBytes(tempPath, workbookBytes);
                sourceWorkbook = application.Workbooks.Open(
                    tempPath,
                    ReadOnly: true,
                    UpdateLinks: 0,
                    AddToMru: false);

                var sourceWorksheet = FindWorksheet(sourceWorkbook, BusinessDataSheetName);
                if (sourceWorksheet == null)
                {
                    throw new InvalidOperationException("The exported workbook does not contain a Business Data sheet.");
                }

                var targetWorksheet = GetWorksheet(targetSheetName);
                var originalTargetSheetName = targetWorksheet.Name;
                CopyBusinessDataSheet(sourceWorkbook, sourceWorksheet, targetWorksheet);
                targetWorksheet.Name = originalTargetSheetName;
            }
            finally
            {
                try
                {
                    sourceWorkbook?.Close(SaveChanges: false);
                }
                catch
                {
                }

                if (File.Exists(tempPath))
                {
                    File.Delete(tempPath);
                }
            }
        }

        public void ActivateWorkSheetAtA1(string sheetName)
        {
            var worksheet = GetWorksheet(sheetName);
            worksheet.Activate();
            var cell = worksheet.Range["A1"] as ExcelInterop.Range;
            cell?.Select();
        }

        private void CopyBusinessDataSheet(
            ExcelInterop.Workbook sourceWorkbook,
            ExcelInterop.Worksheet sourceWorksheet,
            ExcelInterop.Worksheet targetWorksheet)
        {
            var sourceUsedRange = sourceWorksheet.UsedRange;
            targetWorksheet.Cells.Clear();

            if (sourceUsedRange != null)
            {
                var targetStart = targetWorksheet.Range["A1"] as ExcelInterop.Range;
                sourceUsedRange.Copy(targetStart);
                CopyColumnWidths(sourceUsedRange, targetWorksheet);
                CopyRowHeights(sourceUsedRange, targetWorksheet);
            }

            CopyFreezePaneState(sourceWorkbook, sourceWorksheet, targetWorksheet);
        }

        private static void CopyColumnWidths(ExcelInterop.Range sourceUsedRange, ExcelInterop.Worksheet targetWorksheet)
        {
            var firstColumn = sourceUsedRange.Column;
            var columnCount = sourceUsedRange.Columns.Count;
            for (var offset = 0; offset < columnCount; offset++)
            {
                var sourceColumn = sourceUsedRange.Worksheet.Columns[firstColumn + offset] as ExcelInterop.Range;
                var targetColumn = targetWorksheet.Columns[1 + offset] as ExcelInterop.Range;
                if (sourceColumn != null && targetColumn != null)
                {
                    targetColumn.ColumnWidth = sourceColumn.ColumnWidth;
                    targetColumn.Hidden = sourceColumn.Hidden;
                }
            }
        }

        private static void CopyRowHeights(ExcelInterop.Range sourceUsedRange, ExcelInterop.Worksheet targetWorksheet)
        {
            var firstRow = sourceUsedRange.Row;
            var rowCount = sourceUsedRange.Rows.Count;
            for (var offset = 0; offset < rowCount; offset++)
            {
                var sourceRow = sourceUsedRange.Worksheet.Rows[firstRow + offset] as ExcelInterop.Range;
                var targetRow = targetWorksheet.Rows[1 + offset] as ExcelInterop.Range;
                if (sourceRow != null && targetRow != null)
                {
                    targetRow.RowHeight = sourceRow.RowHeight;
                    targetRow.Hidden = sourceRow.Hidden;
                }
            }
        }

        private void CopyFreezePaneState(
            ExcelInterop.Workbook sourceWorkbook,
            ExcelInterop.Worksheet sourceWorksheet,
            ExcelInterop.Worksheet targetWorksheet)
        {
            sourceWorksheet.Activate();
            var sourceWindow = sourceWorkbook.Windows.Count > 0
                ? sourceWorkbook.Windows[1] as ExcelInterop.Window
                : null;
            var splitRow = sourceWindow?.SplitRow ?? 0;
            var splitColumn = sourceWindow?.SplitColumn ?? 0;
            var freezePanes = sourceWindow?.FreezePanes ?? false;

            targetWorksheet.Activate();
            var targetWindow = application.ActiveWindow;
            if (targetWindow == null)
            {
                return;
            }

            targetWindow.FreezePanes = false;
            targetWindow.SplitRow = splitRow;
            targetWindow.SplitColumn = splitColumn;
            targetWindow.FreezePanes = freezePanes;
        }

        private ExcelInterop.Workbook GetWorkbook()
        {
            var workbook = application.ActiveWorkbook;
            if (workbook == null)
            {
                throw new InvalidOperationException("Excel workbook is not available.");
            }

            return workbook;
        }

        private ExcelInterop.Worksheet GetWorksheet(string sheetName)
        {
            var worksheet = FindWorksheet(GetWorkbook(), sheetName);
            if (worksheet != null)
            {
                return worksheet;
            }

            throw new InvalidOperationException($"Worksheet '{sheetName}' was not found.");
        }

        private static ExcelInterop.Worksheet FindWorksheet(ExcelInterop.Workbook workbook, string sheetName)
        {
            if (workbook == null)
            {
                return null;
            }

            for (var index = 1; index <= workbook.Worksheets.Count; index++)
            {
                var worksheet = workbook.Worksheets[index] as ExcelInterop.Worksheet;
                if (worksheet != null &&
                    string.Equals(worksheet.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    return worksheet;
                }
            }

            return null;
        }

        private static bool HasSpecialCells(ExcelInterop.Worksheet worksheet, ExcelInterop.XlCellType cellType)
        {
            try
            {
                var cells = worksheet.Cells.SpecialCells(cellType);
                return cells != null;
            }
            catch (COMException)
            {
                return false;
            }
        }
    }
}
```

- [ ] **Step 5: Include new files in the add-in project**

In `src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj`, add these compile entries near the other `Excel\...` entries:

```xml
<Compile Include="Excel\IBusinessWorkbookImporter.cs" />
<Compile Include="Excel\ExcelBusinessWorkbookImporter.cs" />
```

- [ ] **Step 6: Run importer configuration tests**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~ExcelBusinessWorkbookImporterConfigurationTests"
```

Expected: PASS.

- [ ] **Step 7: Commit importer boundary**

Run:

```powershell
git add src/OfficeAgent.ExcelAddIn/Excel/IBusinessWorkbookImporter.cs src/OfficeAgent.ExcelAddIn/Excel/ExcelBusinessWorkbookImporter.cs src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj tests/OfficeAgent.ExcelAddIn.Tests/ExcelBusinessWorkbookImporterConfigurationTests.cs
git commit -m "feat: import business workbook into worksheet"
```

---

### Task 4: Initialization Dialog And Progress Dialog

**Files:**
- Create: `src/OfficeAgent.ExcelAddIn/Dialogs/InitializeSheetDialogModels.cs`
- Create: `src/OfficeAgent.ExcelAddIn/Dialogs/InitializeSheetDialog.cs`
- Create: `src/OfficeAgent.ExcelAddIn/Dialogs/InitializeSheetImportProgressDialog.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/Dialogs/OperationResultDialog.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/Localization/HostLocalizedStrings.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj`
- Test: `tests/OfficeAgent.ExcelAddIn.Tests/InitializeSheetDialogTests.cs`
- Test: `tests/OfficeAgent.ExcelAddIn.Tests/HostLocalizedStringsTests.cs`

- [ ] **Step 1: Add dialog model tests**

Create `tests/OfficeAgent.ExcelAddIn.Tests/InitializeSheetDialogTests.cs`:

```csharp
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class InitializeSheetDialogTests
    {
        [Theory]
        [InlineData(true, true, "TemplateImport")]
        [InlineData(true, false, "ConfigOnly")]
        [InlineData(false, true, "ConfigOnly")]
        [InlineData(false, false, "ConfigOnly")]
        public void ResolveDefaultModeFollowsBlankSheetPolicy(
            bool isBlankSheet,
            bool canImportTemplate,
            string expectedModeName)
        {
            var type = LoadAddInAssembly().GetType(
                "OfficeAgent.ExcelAddIn.Dialogs.InitializeSheetDialog",
                throwOnError: true);
            var method = type.GetMethod(
                "ResolveDefaultMode",
                BindingFlags.Static | BindingFlags.NonPublic);

            Assert.NotNull(method);
            var mode = method.Invoke(null, new object[] { isBlankSheet, canImportTemplate });
            Assert.Equal(expectedModeName, mode.ToString());
        }

        [Fact]
        public void SourceContainsOverwriteRiskTextAndNoHardcodedDialogCopy()
        {
            var dialogText = ReadSource("src", "OfficeAgent.ExcelAddIn", "Dialogs", "InitializeSheetDialog.cs");
            var stringsText = ReadSource("src", "OfficeAgent.ExcelAddIn", "Localization", "HostLocalizedStrings.cs");

            Assert.Contains("InitializeSheetOverwriteRiskMessage", dialogText, StringComparison.Ordinal);
            Assert.Contains("InitializeSheetOverwriteRiskMessage", stringsText, StringComparison.Ordinal);
            Assert.DoesNotContain("覆盖当前表", dialogText, StringComparison.Ordinal);
        }

        private static Assembly LoadAddInAssembly()
        {
            return Assembly.LoadFrom(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "bin",
                "Debug",
                "OfficeAgent.ExcelAddIn.dll"));
        }

        private static string ReadSource(params string[] segments)
        {
            return File.ReadAllText(ResolveRepositoryPath(segments));
        }

        private static string ResolveRepositoryPath(params string[] segments)
        {
            return Path.GetFullPath(Path.Combine(new[]
            {
                AppContext.BaseDirectory,
                "..",
                "..",
                "..",
                "..",
                "..",
            }.Concat(segments).ToArray()));
        }
    }
}
```

- [ ] **Step 2: Add localization tests**

In `tests/OfficeAgent.ExcelAddIn.Tests/HostLocalizedStringsTests.cs`, add:

```csharp
[Theory]
[InlineData("zh", "初始化当前表", "从模板创建作业表", "仅初始化配置", "覆盖并初始化")]
[InlineData("en", "Initialize sheet", "Create sheet from template", "Initialize configuration only", "Overwrite and initialize")]
public void ForLocaleReturnsInitializeSheetDialogText(
    string locale,
    string expectedTitle,
    string expectedTemplateMode,
    string expectedConfigOnlyMode,
    string expectedOverwriteButton)
{
    var strings = CreateStrings(locale);

    Assert.Equal(expectedTitle, GetString(strings, "InitializeSheetDialogTitle"));
    Assert.Equal(expectedTemplateMode, GetString(strings, "InitializeSheetTemplateModeText"));
    Assert.Equal(expectedConfigOnlyMode, GetString(strings, "InitializeSheetConfigOnlyModeText"));
    Assert.Equal(expectedOverwriteButton, GetString(strings, "InitializeSheetOverwriteButtonText"));
}

[Theory]
[InlineData("zh", "正在下载模板 Excel...", "正在导入到当前工作表...", "正在写入同步配置...")]
[InlineData("en", "Downloading template Excel...", "Importing into the current worksheet...", "Writing sync configuration...")]
public void ForLocaleReturnsInitializeSheetProgressText(
    string locale,
    string expectedDownloading,
    string expectedImporting,
    string expectedWriting)
{
    var strings = CreateStrings(locale);

    Assert.Equal(expectedDownloading, GetString(strings, "InitializeSheetProgressDownloadingText"));
    Assert.Equal(expectedImporting, GetString(strings, "InitializeSheetProgressImportingText"));
    Assert.Equal(expectedWriting, GetString(strings, "InitializeSheetProgressWritingConfigurationText"));
}
```

- [ ] **Step 3: Run dialog/localization tests and verify expected failures**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~InitializeSheetDialogTests|FullyQualifiedName~ForLocaleReturnsInitializeSheetDialogText|FullyQualifiedName~ForLocaleReturnsInitializeSheetProgressText"
```

Expected: FAIL because dialog files and localized string properties do not exist.

- [ ] **Step 4: Add dialog models**

Create `src/OfficeAgent.ExcelAddIn/Dialogs/InitializeSheetDialogModels.cs`:

```csharp
using System;
using System.Collections.Generic;
using OfficeAgent.Core.Models;

namespace OfficeAgent.ExcelAddIn.Dialogs
{
    internal enum InitializeSheetMode
    {
        ConfigOnly,
        TemplateImport,
    }

    internal sealed class InitializeSheetDialogRequest
    {
        public string ProjectDisplayName { get; set; } = string.Empty;

        public bool IsBlankWorkSheet { get; set; }

        public bool SupportsTemplateImport { get; set; }
    }

    internal sealed class InitializeSheetTemplateLoadResult
    {
        public bool IsSupported { get; set; }

        public string DisabledReason { get; set; } = string.Empty;

        public IReadOnlyList<BusinessExportTemplateOption> Templates { get; set; } = Array.Empty<BusinessExportTemplateOption>();

        public static InitializeSheetTemplateLoadResult Unsupported(string reason)
        {
            return new InitializeSheetTemplateLoadResult
            {
                IsSupported = false,
                DisabledReason = reason ?? string.Empty,
            };
        }

        public static InitializeSheetTemplateLoadResult Success(IReadOnlyList<BusinessExportTemplateOption> templates)
        {
            return new InitializeSheetTemplateLoadResult
            {
                IsSupported = true,
                Templates = templates ?? Array.Empty<BusinessExportTemplateOption>(),
            };
        }

        public static InitializeSheetTemplateLoadResult Failed(string reason)
        {
            return new InitializeSheetTemplateLoadResult
            {
                IsSupported = true,
                DisabledReason = reason ?? string.Empty,
            };
        }
    }

    internal sealed class InitializeSheetDialogResult
    {
        public InitializeSheetMode Mode { get; set; }

        public BusinessExportTemplateOption SelectedTemplate { get; set; }
    }
}
```

- [ ] **Step 5: Add localized strings**

In `src/OfficeAgent.ExcelAddIn/Localization/HostLocalizedStrings.cs`, add these properties near `InitializeCurrentSheetCompletedMessage`:

```csharp
public string InitializeSheetDialogTitle => Locale == "zh" ? "初始化当前表" : "Initialize sheet";

public string InitializeSheetTemplateModeText => Locale == "zh" ? "从模板创建作业表" : "Create sheet from template";

public string InitializeSheetConfigOnlyModeText => Locale == "zh" ? "仅初始化配置" : "Initialize configuration only";

public string InitializeSheetTemplateLoadingText => Locale == "zh" ? "正在加载模板..." : "Loading templates...";

public string InitializeSheetTemplateEmptyText => Locale == "zh"
    ? "当前项目暂无可用模板，可先仅初始化配置。"
    : "No templates are available for the current project. You can initialize configuration only.";

public string InitializeSheetTemplateLoadFailedText => Locale == "zh"
    ? "模板加载失败，可先仅初始化配置。"
    : "Templates failed to load. You can initialize configuration only.";

public string InitializeSheetTemplateUnsupportedText => Locale == "zh"
    ? "当前业务系统不支持从模板创建作业表。"
    : "The current business system does not support creating a sheet from a template.";

public string InitializeSheetOverwriteRiskMessage => Locale == "zh"
    ? "当前表已有内容。从模板创建作业表会覆盖当前表内容和格式，请确认你已经备份或不再需要这些内容。"
    : "The current sheet already has content. Creating from a template will overwrite current content and formatting. Confirm you have backed up or no longer need it.";

public string InitializeSheetConfirmButtonText => Locale == "zh" ? "初始化" : "Initialize";

public string InitializeSheetOverwriteButtonText => Locale == "zh" ? "覆盖并初始化" : "Overwrite and initialize";

public string InitializeSheetConfigOnlyCompletedMessage => Locale == "zh"
    ? "初始化完成，当前表内容未修改。你可以继续上传或下载数据。"
    : "Initialization completed. The current sheet content was not changed. You can continue uploading or downloading data.";

public string InitializeSheetTemplateImportCompletedMessage => Locale == "zh"
    ? "初始化完成，已从模板创建当前作业表。你可以编辑数据后上传，或全选需要刷新的区域后点击下载。"
    : "Initialization completed. The current worksheet was created from the template. You can edit data and upload, or select the area to refresh and click Download.";

public string InitializeSheetManagedSheetBlockedMessage => Locale == "zh"
    ? "xISDP_Setting 和 xISDP_Log 不能执行初始化。请切换到业务工作表后重试。"
    : "xISDP_Setting and xISDP_Log cannot be initialized. Switch to a business worksheet and try again.";

public string InitializeSheetMetadataIncompleteMessage => Locale == "zh"
    ? "表内容已导入，但同步配置未完成。请重新初始化当前表。"
    : "Sheet content was imported, but sync configuration was not completed. Reinitialize the current sheet.";

public string InitializeSheetProgressDialogTitle => Locale == "zh" ? "从模板创建作业表" : "Create sheet from template";

public string InitializeSheetProgressDownloadingText => Locale == "zh" ? "正在下载模板 Excel..." : "Downloading template Excel...";

public string InitializeSheetProgressImportingText => Locale == "zh" ? "正在导入到当前工作表..." : "Importing into the current worksheet...";

public string InitializeSheetProgressWritingConfigurationText => Locale == "zh" ? "正在写入同步配置..." : "Writing sync configuration...";

public string InitializeSheetProgressCancelButtonText => Locale == "zh" ? "取消" : "Cancel";
```

- [ ] **Step 6: Implement `InitializeSheetDialog`**

Create `src/OfficeAgent.ExcelAddIn/Dialogs/InitializeSheetDialog.cs` with a fixed dialog shell:

```csharp
using System;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeAgent.Core.Models;
using OfficeAgent.ExcelAddIn.Localization;

namespace OfficeAgent.ExcelAddIn.Dialogs
{
    internal sealed class InitializeSheetDialog : Form
    {
        private readonly InitializeSheetDialogRequest request;
        private readonly Func<InitializeSheetTemplateLoadResult> loadTemplates;
        private readonly HostLocalizedStrings strings;
        private readonly RadioButton templateModeButton;
        private readonly RadioButton configOnlyModeButton;
        private readonly ListBox templateListBox;
        private readonly Label statusLabel;
        private readonly Label riskLabel;
        private readonly Button confirmButton;
        private readonly Button cancelButton;
        private InitializeSheetTemplateLoadResult templateLoadResult;

        public InitializeSheetDialog(
            InitializeSheetDialogRequest request,
            Func<InitializeSheetTemplateLoadResult> loadTemplates,
            HostLocalizedStrings strings = null)
        {
            this.request = request ?? throw new ArgumentNullException(nameof(request));
            this.loadTemplates = loadTemplates ?? throw new ArgumentNullException(nameof(loadTemplates));
            this.strings = strings ?? HostLocalizedStrings.ForLocale("en");

            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Dpi;
            Text = this.strings.InitializeSheetDialogTitle;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(520, 360);
            Padding = new Padding(18);

            templateModeButton = new RadioButton
            {
                Text = this.strings.InitializeSheetTemplateModeText,
                AutoSize = true,
                Location = new Point(18, 18),
            };
            templateModeButton.CheckedChanged += (sender, args) => RefreshState();

            templateListBox = new ListBox
            {
                DisplayMember = nameof(BusinessExportTemplateOption.TemplateName),
                ValueMember = nameof(BusinessExportTemplateOption.TemplateId),
                Location = new Point(38, 48),
                Size = new Size(446, 118),
                Enabled = false,
            };

            statusLabel = new Label
            {
                AutoSize = false,
                Location = new Point(38, 172),
                Size = new Size(446, 42),
                ForeColor = Color.FromArgb(96, 96, 96),
                Text = this.strings.InitializeSheetTemplateLoadingText,
            };

            configOnlyModeButton = new RadioButton
            {
                Text = this.strings.InitializeSheetConfigOnlyModeText,
                AutoSize = true,
                Location = new Point(18, 222),
            };
            configOnlyModeButton.CheckedChanged += (sender, args) => RefreshState();

            riskLabel = new Label
            {
                AutoSize = false,
                Location = new Point(38, 252),
                Size = new Size(446, 46),
                ForeColor = Color.FromArgb(160, 80, 0),
                Text = this.strings.InitializeSheetOverwriteRiskMessage,
                Visible = false,
            };

            confirmButton = new Button
            {
                Text = this.strings.InitializeSheetConfirmButtonText,
                Enabled = false,
                DialogResult = DialogResult.OK,
                Size = new Size(126, 32),
                Location = new Point(276, 312),
            };
            confirmButton.Click += (sender, args) => Result = BuildResult();

            cancelButton = new Button
            {
                Text = this.strings.CancelButtonText,
                DialogResult = DialogResult.Cancel,
                Size = new Size(82, 32),
                Location = new Point(410, 312),
            };

            AcceptButton = confirmButton;
            CancelButton = cancelButton;
            Controls.Add(templateModeButton);
            Controls.Add(templateListBox);
            Controls.Add(statusLabel);
            Controls.Add(configOnlyModeButton);
            Controls.Add(riskLabel);
            Controls.Add(confirmButton);
            Controls.Add(cancelButton);
        }

        public InitializeSheetDialogResult Result { get; private set; }

        internal static InitializeSheetMode ResolveDefaultMode(bool isBlankSheet, bool canImportTemplate)
        {
            return isBlankSheet && canImportTemplate
                ? InitializeSheetMode.TemplateImport
                : InitializeSheetMode.ConfigOnly;
        }

        protected override async void OnShown(EventArgs e)
        {
            base.OnShown(e);
            await LoadTemplatesAsync().ConfigureAwait(true);
        }

        private async Task LoadTemplatesAsync()
        {
            confirmButton.Enabled = false;
            statusLabel.Text = strings.InitializeSheetTemplateLoadingText;
            templateLoadResult = await Task.Run(loadTemplates).ConfigureAwait(true);

            var templates = (templateLoadResult?.Templates ?? Array.Empty<BusinessExportTemplateOption>()).ToArray();
            templateListBox.Items.Clear();
            foreach (var template in templates)
            {
                templateListBox.Items.Add(template);
            }

            if (templateListBox.Items.Count > 0)
            {
                templateListBox.SelectedIndex = 0;
            }

            var canImport = CanImportTemplate();
            var defaultMode = ResolveDefaultMode(request.IsBlankWorkSheet, canImport);
            templateModeButton.Checked = defaultMode == InitializeSheetMode.TemplateImport;
            configOnlyModeButton.Checked = defaultMode == InitializeSheetMode.ConfigOnly;
            RefreshState();
        }

        private bool CanImportTemplate()
        {
            return request.SupportsTemplateImport &&
                   templateLoadResult != null &&
                   templateLoadResult.IsSupported &&
                   templateListBox.Items.Count > 0;
        }

        private void RefreshState()
        {
            var canImport = CanImportTemplate();
            templateModeButton.Enabled = canImport;
            templateListBox.Enabled = canImport && templateModeButton.Checked;
            statusLabel.Text = ResolveStatusText(canImport);
            riskLabel.Visible = templateModeButton.Checked && !request.IsBlankWorkSheet;
            confirmButton.Text = riskLabel.Visible
                ? strings.InitializeSheetOverwriteButtonText
                : strings.InitializeSheetConfirmButtonText;
            confirmButton.Enabled = configOnlyModeButton.Checked || (templateModeButton.Checked && templateListBox.SelectedItem != null);
        }

        private string ResolveStatusText(bool canImport)
        {
            if (canImport)
            {
                return string.Empty;
            }

            if (templateLoadResult == null)
            {
                return strings.InitializeSheetTemplateLoadingText;
            }

            if (!string.IsNullOrWhiteSpace(templateLoadResult.DisabledReason))
            {
                return templateLoadResult.DisabledReason;
            }

            return strings.InitializeSheetTemplateEmptyText;
        }

        private InitializeSheetDialogResult BuildResult()
        {
            return new InitializeSheetDialogResult
            {
                Mode = templateModeButton.Checked ? InitializeSheetMode.TemplateImport : InitializeSheetMode.ConfigOnly,
                SelectedTemplate = templateListBox.SelectedItem as BusinessExportTemplateOption,
            };
        }
    }
}
```

- [ ] **Step 7: Implement progress dialog**

Create `src/OfficeAgent.ExcelAddIn/Dialogs/InitializeSheetImportProgressDialog.cs`:

```csharp
using System;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeAgent.ExcelAddIn.Localization;

namespace OfficeAgent.ExcelAddIn.Dialogs
{
    internal interface IInitializeSheetImportProgress
    {
        void SetDownloading();

        void SetImporting();

        void SetWritingConfiguration();
    }

    internal sealed class InitializeSheetImportProgressDialog : Form, IInitializeSheetImportProgress
    {
        private readonly Func<IInitializeSheetImportProgress, CancellationToken, Task> operation;
        private readonly CancellationTokenSource cancellationTokenSource = new CancellationTokenSource();
        private readonly Label messageLabel;
        private readonly Button cancelButton;
        private readonly HostLocalizedStrings strings;
        private Exception error;
        private bool completed;

        private InitializeSheetImportProgressDialog(
            Func<IInitializeSheetImportProgress, CancellationToken, Task> operation,
            HostLocalizedStrings strings)
        {
            this.operation = operation ?? throw new ArgumentNullException(nameof(operation));
            this.strings = strings ?? HostLocalizedStrings.ForLocale("en");

            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Dpi;
            Text = this.strings.InitializeSheetProgressDialogTitle;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            ControlBox = false;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(410, 150);
            Padding = new Padding(18);

            messageLabel = new Label
            {
                AutoSize = false,
                Dock = DockStyle.Top,
                Height = 58,
                Text = this.strings.InitializeSheetProgressDownloadingText,
                TextAlign = ContentAlignment.MiddleLeft,
            };

            var progressBar = new ProgressBar
            {
                Dock = DockStyle.Top,
                Height = 18,
                MarqueeAnimationSpeed = 30,
                Style = ProgressBarStyle.Marquee,
            };

            cancelButton = new Button
            {
                Text = this.strings.InitializeSheetProgressCancelButtonText,
                AutoSize = true,
                Padding = new Padding(14, 4, 14, 4),
                Anchor = AnchorStyles.Right,
            };
            cancelButton.Click += (sender, args) =>
            {
                cancelButton.Enabled = false;
                cancellationTokenSource.Cancel();
            };

            var buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                FlowDirection = FlowDirection.RightToLeft,
                Height = 42,
                Padding = new Padding(0, 12, 0, 0),
            };
            buttonPanel.Controls.Add(cancelButton);

            Controls.Add(progressBar);
            Controls.Add(messageLabel);
            Controls.Add(buttonPanel);
        }

        public static bool Run(
            IWin32Window owner,
            Func<IInitializeSheetImportProgress, CancellationToken, Task> operation)
        {
            var strings = Globals.ThisAddIn?.HostLocalizedStrings ?? HostLocalizedStrings.ForLocale("en");
            using (var dialog = new InitializeSheetImportProgressDialog(operation, strings))
            {
                var result = owner == null ? dialog.ShowDialog() : dialog.ShowDialog(owner);
                if (dialog.error != null)
                {
                    throw dialog.error;
                }

                return result == DialogResult.OK && dialog.completed;
            }
        }

        public void SetDownloading()
        {
            RunOnUiThread(() =>
            {
                messageLabel.Text = strings.InitializeSheetProgressDownloadingText;
                cancelButton.Enabled = true;
                cancelButton.Visible = true;
            });
        }

        public void SetImporting()
        {
            RunOnUiThread(() =>
            {
                messageLabel.Text = strings.InitializeSheetProgressImportingText;
                cancelButton.Enabled = false;
            });
        }

        public void SetWritingConfiguration()
        {
            RunOnUiThread(() =>
            {
                messageLabel.Text = strings.InitializeSheetProgressWritingConfigurationText;
                cancelButton.Enabled = false;
            });
        }

        protected override async void OnShown(EventArgs e)
        {
            base.OnShown(e);
            try
            {
                await operation(this, cancellationTokenSource.Token).ConfigureAwait(true);
                completed = true;
                DialogResult = DialogResult.OK;
            }
            catch (OperationCanceledException)
            {
                DialogResult = DialogResult.Cancel;
            }
            catch (Exception ex)
            {
                error = ex;
                DialogResult = DialogResult.Abort;
            }
            finally
            {
                Close();
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (!completed)
                {
                    cancellationTokenSource.Cancel();
                }

                cancellationTokenSource.Dispose();
            }

            base.Dispose(disposing);
        }

        private void RunOnUiThread(Action action)
        {
            if (InvokeRequired)
            {
                BeginInvoke(action);
                return;
            }

            action();
        }
    }
}
```

- [ ] **Step 8: Extend `IRibbonSyncDialogService`**

In `src/OfficeAgent.ExcelAddIn/Dialogs/OperationResultDialog.cs`, add these members to `IRibbonSyncDialogService`:

```csharp
InitializeSheetDialogResult ShowInitializeSheetDialog(
    InitializeSheetDialogRequest request,
    Func<InitializeSheetTemplateLoadResult> loadTemplates);

bool RunInitializeSheetTemplateImportWithProgress(
    Func<IInitializeSheetImportProgress, CancellationToken, Task> operation);
```

Add these implementations to `RibbonSyncDialogService`:

```csharp
public InitializeSheetDialogResult ShowInitializeSheetDialog(
    InitializeSheetDialogRequest request,
    Func<InitializeSheetTemplateLoadResult> loadTemplates)
{
    var owner = ExcelDialogOwner.FromCurrentApplication();
    using (var dialog = new InitializeSheetDialog(request, loadTemplates, Globals.ThisAddIn?.HostLocalizedStrings))
    {
        var result = owner == null ? dialog.ShowDialog() : dialog.ShowDialog(owner);
        return result == DialogResult.OK ? dialog.Result : null;
    }
}

public bool RunInitializeSheetTemplateImportWithProgress(
    Func<IInitializeSheetImportProgress, CancellationToken, Task> operation)
{
    var owner = ExcelDialogOwner.FromCurrentApplication();
    return InitializeSheetImportProgressDialog.Run(owner, operation);
}
```

- [ ] **Step 9: Include new dialog files in the add-in project**

In `src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj`, add:

```xml
<Compile Include="Dialogs\InitializeSheetDialogModels.cs" />
<Compile Include="Dialogs\InitializeSheetDialog.cs" />
<Compile Include="Dialogs\InitializeSheetImportProgressDialog.cs" />
```

- [ ] **Step 10: Run dialog and localization tests**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~InitializeSheetDialogTests|FullyQualifiedName~HostLocalizedStringsTests"
```

Expected: PASS.

- [ ] **Step 11: Commit dialog shell**

Run:

```powershell
git add src/OfficeAgent.ExcelAddIn/Dialogs/InitializeSheetDialogModels.cs src/OfficeAgent.ExcelAddIn/Dialogs/InitializeSheetDialog.cs src/OfficeAgent.ExcelAddIn/Dialogs/InitializeSheetImportProgressDialog.cs src/OfficeAgent.ExcelAddIn/Dialogs/OperationResultDialog.cs src/OfficeAgent.ExcelAddIn/Localization/HostLocalizedStrings.cs src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj tests/OfficeAgent.ExcelAddIn.Tests/InitializeSheetDialogTests.cs tests/OfficeAgent.ExcelAddIn.Tests/HostLocalizedStringsTests.cs
git commit -m "feat: add initialize sheet template dialogs"
```

---

### Task 5: Execution Service Template Import Orchestration

**Files:**
- Modify: `src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs`
- Test: `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs`

- [ ] **Step 1: Add orchestration tests**

In `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs`, add:

```csharp
using OfficeAgent.ExcelAddIn.Dialogs;
using OfficeAgent.ExcelAddIn.Excel;

[Fact]
public async Task InitializeCurrentSheetFromBusinessTemplateDownloadsImportsThenWritesMetadata()
{
    var connector = new FakeBusinessTemplateConnector();
    var metadataStore = new FakeWorksheetMetadataStore();
    var importer = new FakeBusinessWorkbookImporter();
    var (service, _) = CreateService(new[] { connector }, metadataStore, new FakeWorksheetSelectionReader(), businessWorkbookImporter: importer);

    await InvokeInitializeFromBusinessTemplateAsync(
        service,
        "Sheet1",
        new ProjectOption
        {
            SystemKey = connector.SystemKey,
            ProjectId = "performance",
            DisplayName = "绩效项目",
        },
        new BusinessExportTemplateOption
        {
            TemplateId = "standard",
            TemplateName = "标准作业表",
        });

    Assert.Equal("performance", connector.LastExportProjectId);
    Assert.Equal("standard", connector.LastExportTemplateId);
    Assert.Equal("Sheet1", importer.ImportedTargetSheetName);
    Assert.NotNull(metadataStore.LastSavedBinding);
    Assert.Equal("Sheet1", metadataStore.LastSavedBinding.SheetName);
    Assert.Equal("A1", importer.LastActivatedAddress);
}

[Fact]
public async Task InitializeCurrentSheetFromBusinessTemplateDoesNotWriteMetadataWhenDownloadFails()
{
    var connector = new FakeBusinessTemplateConnector
    {
        ExportException = new InvalidOperationException("download failed"),
    };
    var metadataStore = new FakeWorksheetMetadataStore();
    var importer = new FakeBusinessWorkbookImporter();
    var (service, _) = CreateService(new[] { connector }, metadataStore, new FakeWorksheetSelectionReader(), businessWorkbookImporter: importer);

    await Assert.ThrowsAsync<InvalidOperationException>(() =>
        InvokeInitializeFromBusinessTemplateAsync(
            service,
            "Sheet1",
            new ProjectOption { SystemKey = connector.SystemKey, ProjectId = "performance", DisplayName = "绩效项目" },
            new BusinessExportTemplateOption { TemplateId = "standard", TemplateName = "标准作业表" }));

    Assert.Null(metadataStore.LastSavedBinding);
    Assert.Empty(metadataStore.LastSavedFieldMappings);
    Assert.Equal(0, importer.ImportCallCount);
}

[Fact]
public async Task InitializeCurrentSheetFromBusinessTemplateThrowsIncompleteMetadataMessageWhenMetadataWriteFailsAfterImport()
{
    var connector = new FakeBusinessTemplateConnector();
    var metadataStore = new FakeWorksheetMetadataStore
    {
        SaveBindingException = new InvalidOperationException("metadata write failed"),
    };
    var importer = new FakeBusinessWorkbookImporter();
    var (service, _) = CreateService(new[] { connector }, metadataStore, new FakeWorksheetSelectionReader(), businessWorkbookImporter: importer);

    var exception = await Assert.ThrowsAsync<InvalidOperationException>(() =>
        InvokeInitializeFromBusinessTemplateAsync(
            service,
            "Sheet1",
            new ProjectOption { SystemKey = connector.SystemKey, ProjectId = "performance", DisplayName = "绩效项目" },
            new BusinessExportTemplateOption { TemplateId = "standard", TemplateName = "标准作业表" }));

    Assert.Contains("sync configuration was not completed", exception.Message);
    Assert.Equal(1, importer.ImportCallCount);
}
```

In the same file, change the existing fake connector declaration from:

```csharp
private sealed class FakeSystemConnector : ISystemConnector, IUploadChangeFilter
```

to:

```csharp
private class FakeSystemConnector : ISystemConnector, IUploadChangeFilter
```

In the same file, add a write-failure hook to `FakeWorksheetMetadataStore`:

```csharp
public Exception SaveBindingException { get; set; }

public void SaveBinding(SheetBinding binding)
{
    if (SaveBindingException != null)
    {
        throw SaveBindingException;
    }

    LastSavedBinding = binding;
    Bindings[binding.SheetName] = binding;
}
```

Replace the `CreateService` overload that accepts `IReadOnlyList<FakeSystemConnector>` with this version:

```csharp
private static (object Service, FakeWorksheetGridAdapter Grid) CreateService(
    IReadOnlyList<FakeSystemConnector> connectors,
    IWorksheetMetadataStore metadataStore,
    FakeWorksheetSelectionReader selectionReader,
    IAiColumnMappingClient aiClient = null,
    IBusinessWorkbookImporter businessWorkbookImporter = null)
{
    var assembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
    var serviceType = assembly.GetType("OfficeAgent.ExcelAddIn.WorksheetSyncExecutionService", throwOnError: true);
    var gridInterface = assembly.GetType("OfficeAgent.ExcelAddIn.Excel.IWorksheetGridAdapter", throwOnError: true);
    var grid = new FakeWorksheetGridAdapter(gridInterface);
    var syncService = new WorksheetSyncService(
        new SystemConnectorRegistry(connectors.Cast<ISystemConnector>().ToArray()),
        metadataStore,
        new WorksheetChangeTracker(),
        new SyncOperationPreviewFactory());

    if (businessWorkbookImporter != null)
    {
        var fullCtor = serviceType.GetConstructor(
            BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
            binder: null,
            types: new[]
            {
                typeof(WorksheetSyncService),
                typeof(IWorksheetMetadataStore),
                typeof(IWorksheetSelectionReader),
                gridInterface,
                typeof(SyncOperationPreviewFactory),
                typeof(IWorksheetChangeLogStore),
                typeof(WorksheetPendingEditTracker),
                typeof(IAiColumnMappingClient),
                typeof(IBusinessWorkbookImporter),
            },
            modifiers: null);

        if (fullCtor == null)
        {
            throw new InvalidOperationException("WorksheetSyncExecutionService template import constructor was not found.");
        }

        var service = fullCtor.Invoke(new object[]
        {
            syncService,
            metadataStore,
            selectionReader,
            grid.GetTransparentProxy(),
            new SyncOperationPreviewFactory(),
            null,
            null,
            aiClient ?? new FakeAiColumnMappingClient(),
            businessWorkbookImporter,
        });

        return (service, grid);
    }

    var ctor = serviceType.GetConstructor(
        BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
        binder: null,
        types: new[]
        {
            typeof(WorksheetSyncService),
            typeof(IWorksheetMetadataStore),
            typeof(IWorksheetSelectionReader),
            gridInterface,
            typeof(SyncOperationPreviewFactory),
        },
        modifiers: null);

    if (ctor == null)
    {
        throw new InvalidOperationException("WorksheetSyncExecutionService constructor was not found.");
    }

    var defaultService = ctor.Invoke(new object[]
    {
        syncService,
        metadataStore,
        selectionReader,
        grid.GetTransparentProxy(),
        new SyncOperationPreviewFactory(),
    });

    if (aiClient == null)
    {
        return (defaultService, grid);
    }

    var aiCtor = serviceType.GetConstructor(
        BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
        binder: null,
        types: new[]
        {
            typeof(WorksheetSyncService),
            typeof(IWorksheetMetadataStore),
            typeof(IWorksheetSelectionReader),
            gridInterface,
            typeof(SyncOperationPreviewFactory),
            typeof(IAiColumnMappingClient),
        },
        modifiers: null);

    if (aiCtor == null)
    {
        throw new InvalidOperationException("WorksheetSyncExecutionService AI constructor was not found.");
    }

    var aiService = aiCtor.Invoke(new object[]
    {
        syncService,
        metadataStore,
        selectionReader,
        grid.GetTransparentProxy(),
        new SyncOperationPreviewFactory(),
        aiClient,
    });

    return (aiService, grid);
}
```

Add this reflection helper near `InvokeInitialize`:

```csharp
private static async Task InvokeInitializeFromBusinessTemplateAsync(
    object service,
    string sheetName,
    ProjectOption project,
    BusinessExportTemplateOption template)
{
    var progress = new FakeInitializeSheetImportProgress();
    var method = service.GetType().GetMethod(
        "InitializeCurrentSheetFromBusinessTemplateAsync",
        BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

    if (method == null)
    {
        throw new InvalidOperationException("InitializeCurrentSheetFromBusinessTemplateAsync was not found.");
    }

    var task = (Task)method.Invoke(service, new object[]
    {
        sheetName,
        project,
        template,
        progress,
        CancellationToken.None,
    });

    await task.ConfigureAwait(false);
}
```

Add this fake importer inside the test class:

```csharp
private sealed class FakeBusinessWorkbookImporter : IBusinessWorkbookImporter
{
    public bool IsBlank { get; set; } = true;

    public int ImportCallCount { get; private set; }

    public string ImportedTargetSheetName { get; private set; }

    public string LastActivatedAddress { get; private set; }

    public bool IsWorkSheetContentBlank(string sheetName)
    {
        return IsBlank;
    }

    public void EnsureCanWriteToWorkSheet(string sheetName)
    {
    }

    public void ImportBusinessDataSheet(byte[] workbookBytes, string targetSheetName)
    {
        ImportCallCount++;
        ImportedTargetSheetName = targetSheetName;
    }

    public void ActivateWorkSheetAtA1(string sheetName)
    {
        LastActivatedAddress = "A1";
    }
}
```

Add this fake progress class inside the test class:

```csharp
private sealed class FakeInitializeSheetImportProgress : IInitializeSheetImportProgress
{
    public void SetDownloading()
    {
    }

    public void SetImporting()
    {
    }

    public void SetWritingConfiguration()
    {
    }
}
```

Add fake business connector:

```csharp
private sealed class FakeBusinessTemplateConnector : FakeSystemConnector, IBusinessExportTemplateConnector
{
    public Exception ExportException { get; set; }

    public string LastExportProjectId { get; private set; }

    public string LastExportTemplateId { get; private set; }

    public IReadOnlyList<BusinessExportTemplateOption> GetBusinessExportTemplates(string projectId)
    {
        return new[]
        {
            new BusinessExportTemplateOption { TemplateId = "standard", TemplateName = "标准作业表" },
        };
    }

    public Task<BusinessExportWorkbook> ExportBusinessWorkbookAsync(
        string projectId,
        string templateId,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (ExportException != null)
        {
            throw ExportException;
        }

        LastExportProjectId = projectId;
        LastExportTemplateId = templateId;
        return Task.FromResult(new BusinessExportWorkbook
        {
            Content = new byte[] { 0x50, 0x4B, 0x03, 0x04 },
        });
    }
}
```

- [ ] **Step 2: Run execution tests and verify expected failures**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~InitializeCurrentSheetFromBusinessTemplate"
```

Expected: FAIL because `WorksheetSyncExecutionService` has no template import method or importer constructor.

- [ ] **Step 3: Add importer dependency to `WorksheetSyncExecutionService`**

In `src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs`, add field:

```csharp
private readonly IBusinessWorkbookImporter businessWorkbookImporter;
```

Add this constructor overload:

```csharp
public WorksheetSyncExecutionService(
    WorksheetSyncService worksheetSyncService,
    IWorksheetMetadataStore metadataStore,
    IWorksheetSelectionReader selectionReader,
    IWorksheetGridAdapter gridAdapter,
    SyncOperationPreviewFactory previewFactory,
    IWorksheetChangeLogStore changeLogStore,
    WorksheetPendingEditTracker pendingEditTracker,
    IAiColumnMappingClient aiColumnMappingClient,
    IBusinessWorkbookImporter businessWorkbookImporter)
    : this(
        worksheetSyncService,
        metadataStore,
        selectionReader,
        gridAdapter,
        previewFactory,
        changeLogStore,
        pendingEditTracker,
        aiColumnMappingClient)
{
    this.businessWorkbookImporter = businessWorkbookImporter ?? throw new ArgumentNullException(nameof(businessWorkbookImporter));
}
```

At the end of the existing main constructor that takes `changeLogStore` and `pendingEditTracker`, add:

```csharp
businessWorkbookImporter = null;
```

- [ ] **Step 4: Add template list and blank-sheet helpers**

Add these public methods to `WorksheetSyncExecutionService`:

```csharp
public bool IsWorkSheetContentBlank(string sheetName)
{
    if (businessWorkbookImporter == null)
    {
        return false;
    }

    return businessWorkbookImporter.IsWorkSheetContentBlank(sheetName);
}

public bool SupportsBusinessExportTemplates(string systemKey)
{
    return worksheetSyncService.SupportsBusinessExportTemplates(systemKey);
}

public InitializeSheetTemplateLoadResult LoadBusinessExportTemplates(ProjectOption project)
{
    var strings = GetStrings();
    if (project == null || string.IsNullOrWhiteSpace(project.SystemKey))
    {
        return InitializeSheetTemplateLoadResult.Unsupported(strings.InitializeSheetTemplateUnsupportedText);
    }

    if (!worksheetSyncService.SupportsBusinessExportTemplates(project.SystemKey))
    {
        return InitializeSheetTemplateLoadResult.Unsupported(strings.InitializeSheetTemplateUnsupportedText);
    }

    try
    {
        var templates = worksheetSyncService.GetBusinessExportTemplates(project.SystemKey, project.ProjectId);
        if (templates == null || templates.Count == 0)
        {
            return InitializeSheetTemplateLoadResult.Failed(strings.InitializeSheetTemplateEmptyText);
        }

        return InitializeSheetTemplateLoadResult.Success(templates);
    }
    catch
    {
        return InitializeSheetTemplateLoadResult.Failed(strings.InitializeSheetTemplateLoadFailedText);
    }
}
```

Add `using OfficeAgent.ExcelAddIn.Dialogs;` to the file.

- [ ] **Step 5: Add template import method**

Add this method to `WorksheetSyncExecutionService`:

```csharp
public async Task InitializeCurrentSheetFromBusinessTemplateAsync(
    string sheetName,
    ProjectOption project,
    BusinessExportTemplateOption template,
    IInitializeSheetImportProgress progress,
    CancellationToken cancellationToken)
{
    if (businessWorkbookImporter == null)
    {
        throw new InvalidOperationException(GetStrings().InitializeSheetTemplateUnsupportedText);
    }

    if (template == null || string.IsNullOrWhiteSpace(template.TemplateId))
    {
        throw new InvalidOperationException(GetStrings().TemplatePickerSelectionRequiredMessage);
    }

    businessWorkbookImporter.EnsureCanWriteToWorkSheet(sheetName);

    progress?.SetDownloading();
    var workbook = await worksheetSyncService.ExportBusinessWorkbookAsync(
        project.SystemKey,
        project.ProjectId,
        template.TemplateId,
        cancellationToken).ConfigureAwait(true);

    cancellationToken.ThrowIfCancellationRequested();
    var initializationPlan = worksheetSyncService.PrepareSheetInitialization(sheetName, project);

    progress?.SetImporting();
    businessWorkbookImporter.ImportBusinessDataSheet(workbook.Content, sheetName);

    try
    {
        progress?.SetWritingConfiguration();
        worksheetSyncService.SaveSheetInitialization(initializationPlan);
    }
    catch (Exception ex)
    {
        throw new InvalidOperationException(GetStrings().InitializeSheetMetadataIncompleteMessage, ex);
    }

    businessWorkbookImporter.ActivateWorkSheetAtA1(sheetName);
}
```

- [ ] **Step 6: Wire importer in `ThisAddIn`**

In `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs`, change the `WorksheetSyncExecutionService` construction to pass a new importer:

```csharp
WorksheetSyncExecutionService = new WorksheetSyncExecutionService(
    WorksheetSyncService,
    WorksheetMetadataStore,
    new ExcelVisibleSelectionReader(Application),
    worksheetGridAdapter,
    new SyncOperationPreviewFactory(),
    WorksheetChangeLogStore,
    WorksheetPendingEditTracker,
    new AiColumnMappingClient(SettingsStore),
    new ExcelBusinessWorkbookImporter(Application));
```

- [ ] **Step 7: Run execution service tests**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~WorksheetSyncExecutionServiceTests"
```

Expected: PASS.

- [ ] **Step 8: Commit execution orchestration**

Run:

```powershell
git add src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs src/OfficeAgent.ExcelAddIn/ThisAddIn.cs tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs
git commit -m "feat: orchestrate template-based sheet initialization"
```

---

### Task 6: Ribbon Initialization Flow, Managed Sheet Guard, And Analytics

**Files:**
- Modify: `src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs`
- Modify: `tests/OfficeAgent.ExcelAddIn.Tests/RibbonSyncControllerTests.cs`

- [ ] **Step 1: Add Ribbon controller tests**

In `tests/OfficeAgent.ExcelAddIn.Tests/RibbonSyncControllerTests.cs`, add:

```csharp
using OfficeAgent.ExcelAddIn.Dialogs;
using OfficeAgent.ExcelAddIn.Excel;

[Fact]
public void ExecuteInitializeCurrentSheetShowsDialogAndRunsConfigOnlyPath()
{
    var connector = new FakeBusinessTemplateConnector();
    var metadataStore = new FakeWorksheetMetadataStore();
    var dialogService = new FakeDialogService
    {
        InitializeSheetDialogResult = CreateConfigOnlyInitializeResult(),
    };
    metadataStore.Bindings["Sheet1"] = CreateBinding("Sheet1", "performance", "绩效项目");
    var controller = CreateController(connector, metadataStore, dialogService, () => "Sheet1");
    InvokeRefresh(controller);

    InvokeExecuteInitializeCurrentSheet(controller);

    Assert.Single(dialogService.InitializeSheetDialogRequests);
    Assert.Equal("performance", connector.LastBuildFieldMappingSeedProjectId);
    Assert.Contains(dialogService.InfoMessages, message => message.Contains("content was not changed"));
}

[Fact]
public void ExecuteInitializeCurrentSheetRunsTemplatePathThroughProgressDialog()
{
    var connector = new FakeBusinessTemplateConnector();
    var metadataStore = new FakeWorksheetMetadataStore();
    var dialogService = new FakeDialogService
    {
        InitializeSheetDialogResult = CreateTemplateInitializeResult(),
    };
    metadataStore.Bindings["Sheet1"] = CreateBinding("Sheet1", "performance", "绩效项目");
    var controller = CreateController(connector, metadataStore, dialogService, () => "Sheet1");
    InvokeRefresh(controller);

    InvokeExecuteInitializeCurrentSheet(controller);

    Assert.Equal(1, dialogService.InitializeTemplateProgressRunCount);
    Assert.Contains(dialogService.InfoMessages, message => message.Contains("created from the template"));
}

[Theory]
[InlineData("xISDP_Setting")]
[InlineData("xISDP_Log")]
public void ExecuteInitializeCurrentSheetBlocksManagedSheets(string sheetName)
{
    var connector = new FakeBusinessTemplateConnector();
    var dialogService = new FakeDialogService();
    var controller = CreateController(connector, new FakeWorksheetMetadataStore(), dialogService, () => sheetName);
    InvokeSelectProject(controller, new ProjectOption
    {
        SystemKey = "current-business-system",
        ProjectId = "performance",
        DisplayName = "绩效项目",
    });

    InvokeExecuteInitializeCurrentSheet(controller);

    Assert.Empty(dialogService.InitializeSheetDialogRequests);
    Assert.Single(dialogService.WarningMessages);
}

[Fact]
public void ExecuteInitializeCurrentSheetTemplateCancellationDoesNotShowError()
{
    var connector = new FakeBusinessTemplateConnector();
    var metadataStore = new FakeWorksheetMetadataStore();
    var dialogService = new FakeDialogService
    {
        InitializeSheetDialogResult = CreateTemplateInitializeResult(),
        CancelInitializeTemplateProgress = true,
    };
    metadataStore.Bindings["Sheet1"] = CreateBinding("Sheet1", "performance", "绩效项目");
    var controller = CreateController(connector, metadataStore, dialogService, () => "Sheet1");
    InvokeRefresh(controller);

    InvokeExecuteInitializeCurrentSheet(controller);

    Assert.Empty(dialogService.ErrorMessages);
    Assert.Empty(dialogService.InfoMessages);
}
```

In the same file, change the existing fake connector declaration from:

```csharp
private sealed class FakeSystemConnector : ISystemConnector, IUploadChangeFilter
```

to:

```csharp
private class FakeSystemConnector : ISystemConnector, IUploadChangeFilter
```

Add helpers near the other helper methods:

```csharp
private static SheetBinding CreateBinding(string sheetName, string projectId, string projectName)
{
    return new SheetBinding
    {
        SheetName = sheetName,
        SystemKey = "current-business-system",
        ProjectId = projectId,
        ProjectName = projectName,
        HeaderStartRow = 1,
        HeaderRowCount = 2,
        DataStartRow = 3,
    };
}

private static InitializeSheetDialogResult CreateConfigOnlyInitializeResult()
{
    return new InitializeSheetDialogResult
    {
        Mode = InitializeSheetMode.ConfigOnly,
    };
}

private static InitializeSheetDialogResult CreateTemplateInitializeResult()
{
    return new InitializeSheetDialogResult
    {
        Mode = InitializeSheetMode.TemplateImport,
        SelectedTemplate = new BusinessExportTemplateOption
        {
            TemplateId = "standard",
            TemplateName = "标准作业表",
        },
    };
}
```

Add this business template fake connector inside the test class:

```csharp
private sealed class FakeBusinessTemplateConnector : FakeSystemConnector, IBusinessExportTemplateConnector
{
    public string LastExportProjectId { get; private set; }

    public string LastExportTemplateId { get; private set; }

    public IReadOnlyList<BusinessExportTemplateOption> GetBusinessExportTemplates(string projectId)
    {
        return new[]
        {
            new BusinessExportTemplateOption
            {
                TemplateId = "standard",
                TemplateName = "标准作业表",
            },
        };
    }

    public Task<BusinessExportWorkbook> ExportBusinessWorkbookAsync(
        string projectId,
        string templateId,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        LastExportProjectId = projectId;
        LastExportTemplateId = templateId;
        return Task.FromResult(new BusinessExportWorkbook
        {
            Content = new byte[] { 0x50, 0x4B, 0x03, 0x04 },
        });
    }
}
```

Add this importer fake inside the test class:

```csharp
private sealed class FakeBusinessWorkbookImporter : IBusinessWorkbookImporter
{
    public bool IsBlank { get; set; } = true;

    public int ImportCallCount { get; private set; }

    public bool IsWorkSheetContentBlank(string sheetName)
    {
        return IsBlank;
    }

    public void EnsureCanWriteToWorkSheet(string sheetName)
    {
    }

    public void ImportBusinessDataSheet(byte[] workbookBytes, string targetSheetName)
    {
        ImportCallCount++;
    }

    public void ActivateWorkSheetAtA1(string sheetName)
    {
    }
}
```

Modify `CreateController` so it passes a template-capable importer into `CreateExecutionService`. Replace:

```csharp
var (executionService, _) = CreateExecutionService(addInAssembly, connector, metadataStore);
```

with:

```csharp
var templateImporter = connector is IBusinessExportTemplateConnector
    ? new FakeBusinessWorkbookImporter()
    : null;
var (executionService, _) = CreateExecutionService(
    addInAssembly,
    connector,
    metadataStore,
    businessWorkbookImporter: templateImporter);
```

Change the `CreateExecutionService` signature to:

```csharp
private static (object Service, FakeWorksheetGridAdapter Grid) CreateExecutionService(
    Assembly addInAssembly,
    FakeSystemConnector connector,
    FakeWorksheetMetadataStore metadataStore,
    IAiColumnMappingClient aiClient = null,
    FakeWorksheetSelectionReader selectionReader = null,
    IBusinessWorkbookImporter businessWorkbookImporter = null)
```

Inside `CreateExecutionService`, after creating `syncService`, add this branch before the existing constructor lookup:

```csharp
if (businessWorkbookImporter != null)
{
    var fullCtor = serviceType.GetConstructor(
        BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
        binder: null,
        types: new[]
        {
            typeof(WorksheetSyncService),
            typeof(IWorksheetMetadataStore),
            typeof(IWorksheetSelectionReader),
            gridInterface,
            typeof(SyncOperationPreviewFactory),
            typeof(IWorksheetChangeLogStore),
            typeof(WorksheetPendingEditTracker),
            typeof(IAiColumnMappingClient),
            typeof(IBusinessWorkbookImporter),
        },
        modifiers: null);

    if (fullCtor == null)
    {
        throw new InvalidOperationException("WorksheetSyncExecutionService template import constructor was not found.");
    }

    var service = fullCtor.Invoke(new object[]
    {
        syncService,
        metadataStore,
        selectionReader ?? new FakeWorksheetSelectionReader(),
        grid.GetTransparentProxy(),
        new SyncOperationPreviewFactory(),
        null,
        null,
        aiClient ?? new FakeAiColumnMappingClient(),
        businessWorkbookImporter,
    });

    return (service, grid);
}
```

Extend `FakeDialogService.Invoke` with:

```csharp
case "ShowInitializeSheetDialog":
    InitializeSheetDialogRequests.Add((InitializeSheetDialogRequest)call.InArgs[0]);
    var loadTemplates = (Func<InitializeSheetTemplateLoadResult>)call.InArgs[1];
    LastTemplateLoadResult = loadTemplates();
    return new ReturnMessage(InitializeSheetDialogResult, null, 0, call.LogicalCallContext, call);
case "RunInitializeSheetTemplateImportWithProgress":
    InitializeTemplateProgressRunCount++;
    var operation = (Func<IInitializeSheetImportProgress, CancellationToken, Task>)call.InArgs[0];
    using (var cancellationTokenSource = new CancellationTokenSource())
    {
        if (CancelInitializeTemplateProgress)
        {
            cancellationTokenSource.Cancel();
            return new ReturnMessage(false, null, 0, call.LogicalCallContext, call);
        }

        operation(new FakeInitializeSheetImportProgress(), cancellationTokenSource.Token).GetAwaiter().GetResult();
        return new ReturnMessage(true, null, 0, call.LogicalCallContext, call);
    }
```

Add properties to `FakeDialogService`:

```csharp
public List<InitializeSheetDialogRequest> InitializeSheetDialogRequests { get; } = new List<InitializeSheetDialogRequest>();

public InitializeSheetTemplateLoadResult LastTemplateLoadResult { get; private set; }

public InitializeSheetDialogResult InitializeSheetDialogResult { get; set; }

public int InitializeTemplateProgressRunCount { get; private set; }

public bool CancelInitializeTemplateProgress { get; set; }
```

Add `FakeInitializeSheetImportProgress`:

```csharp
private sealed class FakeInitializeSheetImportProgress : IInitializeSheetImportProgress
{
    public void SetDownloading()
    {
    }

    public void SetImporting()
    {
    }

    public void SetWritingConfiguration()
    {
    }
}
```

- [ ] **Step 2: Run Ribbon tests and verify expected failures**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~ExecuteInitializeCurrentSheet"
```

Expected: FAIL because `ExecuteInitializeCurrentSheet` still directly initializes without dialog.

- [ ] **Step 3: Add managed sheet guard**

In `src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs`, add:

```csharp
private static bool IsBlockedInitializationSheet(string sheetName)
{
    return MetadataWorksheetNames.IsMetadataWorksheet(sheetName) ||
           string.Equals(sheetName, "xISDP_Log", StringComparison.OrdinalIgnoreCase);
}
```

At the start of `ExecuteInitializeCurrentSheet`, after `sheetName = GetRequiredSheetName();`, add:

```csharp
if (IsBlockedInitializationSheet(sheetName))
{
    dialogService.ShowWarning(GetStrings().InitializeSheetManagedSheetBlockedMessage);
    return;
}
```

- [ ] **Step 4: Replace direct initialization with dialog flow**

Replace the success body in `ExecuteInitializeCurrentSheet` with:

```csharp
var service = EnsureExecutionService();
var isBlankWorkSheet = service.IsWorkSheetContentBlank(sheetName);
var request = new InitializeSheetDialogRequest
{
    ProjectDisplayName = project.DisplayName ?? string.Empty,
    IsBlankWorkSheet = isBlankWorkSheet,
    SupportsTemplateImport = service.SupportsBusinessExportTemplates(project.SystemKey),
};

var choice = dialogService.ShowInitializeSheetDialog(
    request,
    () => service.LoadBusinessExportTemplates(project));
if (choice == null)
{
    TrackRibbonEvent("ribbon.initialize.canceled");
    return;
}

if (choice.Mode == InitializeSheetMode.ConfigOnly)
{
    service.InitializeCurrentSheet(sheetName, project);
    OfficeAgentLog.Info(
        "ribbon_sync",
        "initialize_sheet.completed",
        "Current worksheet initialized without modifying business cells.",
        BuildInitializeSheetDetails(sheetName, project));
    TrackRibbonEvent("ribbon.initialize.completed");
    dialogService.ShowInfo(GetStrings().InitializeSheetConfigOnlyCompletedMessage);
    return;
}

TrackRibbonEvent(
    "ribbon.initialize_template_import.started",
    new Dictionary<string, object>(StringComparer.Ordinal)
    {
        ["templateId"] = choice.SelectedTemplate?.TemplateId ?? string.Empty,
        ["isBlankSheet"] = isBlankWorkSheet,
    });

var completed = dialogService.RunInitializeSheetTemplateImportWithProgress(
    (progress, cancellationToken) => service.InitializeCurrentSheetFromBusinessTemplateAsync(
        sheetName,
        project,
        choice.SelectedTemplate,
        progress,
        cancellationToken));
if (!completed)
{
    TrackRibbonEvent(
        "ribbon.initialize_template_import.canceled",
        new Dictionary<string, object>(StringComparer.Ordinal)
        {
            ["templateId"] = choice.SelectedTemplate?.TemplateId ?? string.Empty,
            ["isBlankSheet"] = isBlankWorkSheet,
        });
    OfficeAgentLog.Info(
        "ribbon_sync",
        "initialize_template_import.canceled",
        "Template workbook download was canceled.",
        BuildInitializeSheetDetails(sheetName, project));
    return;
}

TrackRibbonEvent(
    "ribbon.initialize_template_import.completed",
    new Dictionary<string, object>(StringComparer.Ordinal)
    {
        ["templateId"] = choice.SelectedTemplate?.TemplateId ?? string.Empty,
        ["isBlankSheet"] = isBlankWorkSheet,
    });
dialogService.ShowInfo(GetStrings().InitializeSheetTemplateImportCompletedMessage);
```

Ensure the existing `AuthenticationRequiredException` catch still calls `HandleAuthenticationRequired(ex)`. In the generic catch, also track template failures when the selected choice was template mode by adding `failedStage = "initialize_template_import"` to the event properties if that local variable is available.

- [ ] **Step 5: Run Ribbon tests**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~RibbonSyncControllerTests"
```

Expected: PASS.

- [ ] **Step 6: Commit Ribbon flow**

Run:

```powershell
git add src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs tests/OfficeAgent.ExcelAddIn.Tests/RibbonSyncControllerTests.cs
git commit -m "feat: route initialization through template dialog"
```

---

### Task 7: Mock Server Template Export

**Files:**
- Modify: `tests/mock-server/package.json`
- Modify: `tests/mock-server/package-lock.json`
- Modify: `tests/mock-server/server.js`
- Modify: `tests/mock-server/README.md`

- [ ] **Step 1: Install the workbook writer dependency**

Run:

```powershell
cd tests/mock-server
npm install xlsx@0.18.5
cd ..\..
```

Expected: `tests/mock-server/package.json` and `tests/mock-server/package-lock.json` include `xlsx`.

- [ ] **Step 2: Add mock business templates**

In `tests/mock-server/server.js`, add near the other top-level requires:

```javascript
const XLSX = require("xlsx");
```

Add after `connectorProjects`:

```javascript
const connectorProjectTemplates = Object.keys(connectorProjectData).reduce(function (acc, projectId) {
  acc[projectId] = [
    { templateId: "standard", templateName: "标准作业表" },
    { templateId: "review", templateName: "评审作业表" },
  ];
  return acc;
}, {});
```

Add workbook builder helpers before the API route section:

```javascript
function createBusinessExportWorkbook(project, templateId) {
  var workbook = XLSX.utils.book_new();
  var headRow1 = ["ID", "负责人"];
  var headRow2 = ["", ""];
  var activityHeads = project.headList.filter(function (head) { return head.headType === "activity"; });
  activityHeads.forEach(function (activity) {
    headRow1.push(activity.activityName, activity.activityName);
    headRow2.push("开始时间", "结束时间");
  });

  var dataRows = project.rows.slice(0, 20).map(function (row) {
    var values = [row.row_id, row.owner_name];
    activityHeads.forEach(function (activity) {
      values.push(row["start_" + activity.activityId] || "", row["end_" + activity.activityId] || "");
    });
    return values;
  });

  var sheet = XLSX.utils.aoa_to_sheet([
    headRow1,
    headRow2,
  ].concat(dataRows));
  sheet["!cols"] = headRow1.map(function () { return { wch: 18 }; });
  sheet["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 1, c: 0 } },
    { s: { r: 0, c: 1 }, e: { r: 1, c: 1 } },
  ];
  activityHeads.forEach(function (_activity, index) {
    var startColumn = 2 + (index * 2);
    sheet["!merges"].push({ s: { r: 0, c: startColumn }, e: { r: 0, c: startColumn + 1 } });
  });

  XLSX.utils.book_append_sheet(workbook, sheet, "Business Data");
  workbook.Props = {
    Title: "OfficeAgent mock business export " + templateId,
    CreatedDate: new Date(Date.UTC(2026, 5, 25)),
  };
  return XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });
}
```

- [ ] **Step 3: Add template endpoints**

Add these routes before `/head`:

```javascript
apiApp.post("/templates", requireAuth, function (req, res) {
  var project = resolveConnectorProject((req.body || {}).projectId, res);
  if (!project) {
    return;
  }

  res.json(connectorProjectTemplates[project.projectId] || []);
});

apiApp.post("/export", requireAuth, function (req, res) {
  var body = req.body || {};
  var project = resolveConnectorProject(body.projectId, res);
  if (!project) {
    return;
  }

  var templates = connectorProjectTemplates[project.projectId] || [];
  var template = templates.find(function (item) { return item.templateId === body.templateId; });
  if (!template) {
    return res.status(404).json({ code: "not_found", message: "未找到模板。" });
  }

  var buffer = createBusinessExportWorkbook(project, template.templateId);
  res
    .type("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    .set("Content-Disposition", "attachment; filename=\"business-export.xlsx\"")
    .send(buffer);
});
```

Add startup logs:

```javascript
console.log("  Template list         = http://localhost:3200/templates");
console.log("  Template export       = http://localhost:3200/export");
```

- [ ] **Step 4: Document mock endpoints**

In `tests/mock-server/README.md`, under Ribbon Sync endpoints, add:

```markdown
#### `POST /templates`

用于：

- 初始化当前表时加载业务系统模板列表

请求体：

```json
{ "projectId": "performance" }
```

返回：

```json
[
  { "templateId": "standard", "templateName": "标准作业表" },
  { "templateId": "review", "templateName": "评审作业表" }
]
```

#### `POST /export`

用于：

- 按 `projectId + templateId` 导出包含 `Business Data` sheet 的 `.xlsx` 二进制文件

请求体：

```json
{ "projectId": "performance", "templateId": "standard" }
```

返回：

- `Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`
- workbook 内包含名为 `Business Data` 的 sheet
```

- [ ] **Step 5: Start mock server and smoke-test endpoints**

Run:

```powershell
cd tests/mock-server
npm start
```

In another terminal, after signing in through the add-in or browser, use the plugin for full validation. For command-only smoke testing without cookies, temporarily inspect the server logs and route registration; the endpoints are protected by `requireAuth` and should return `401` without a session cookie.

- [ ] **Step 6: Commit mock server**

Run:

```powershell
git add tests/mock-server/package.json tests/mock-server/package-lock.json tests/mock-server/server.js tests/mock-server/README.md
git commit -m "test: mock business template export"
```

---

### Task 8: Behavior Docs, Manual Checklist, And Full Verification

**Files:**
- Modify: `docs/modules/ribbon-sync-current-behavior.md`
- Modify: `docs/ribbon-sync-real-system-integration-guide.md`
- Modify: `docs/vsto-manual-test-checklist.md`
- Modify: `docs/module-index.md`

- [ ] **Step 1: Update Ribbon Sync current behavior**

In `docs/modules/ribbon-sync-current-behavior.md`, update section `4.2 显式初始化` to describe:

```markdown
`初始化当前表` 现在总是先打开初始化对话框。

- 空白业务工作表且业务系统模板可用时，默认选择 `从模板创建作业表`
- 非空业务工作表默认选择 `仅初始化配置`
- 用户可以在非空业务工作表上手动选择模板导入，但弹窗会提示覆盖风险，主按钮显示 `覆盖并初始化`
- `仅初始化配置` 保持旧语义，只刷新 `SheetBindings + SheetFieldMappings`，不改动业务单元格
- `从模板创建作业表` 会下载业务系统导出的 `.xlsx`，复制其中 `Business Data` sheet 到当前工作表，保留当前工作表名称，然后写入同步配置
- `xISDP_Setting` 和 `xISDP_Log` 不能执行初始化
- 模板下载阶段可以取消；取消后不写 Setting，不改当前表，不显示错误
```

Also update section `8 当前业务系统合同` with:

```markdown
当前系统还实现可选扩展 `IBusinessExportTemplateConnector`：

- `POST /templates`：按 `projectId` 返回 `{ templateId, templateName }[]`
- `POST /export`：按 `projectId + templateId` 返回 `.xlsx` 二进制
- 导出的 workbook 必须包含 `Business Data` sheet
- 插件不要求额外字段映射 metadata，初始化配置仍由 `BuildFieldMappingSeed` 生成
```

- [ ] **Step 2: Update real-system integration guide**

In `docs/ribbon-sync-real-system-integration-guide.md`, add a section titled `业务导出模板扩展` with this code block:

```csharp
public interface IBusinessExportTemplateConnector
{
    IReadOnlyList<BusinessExportTemplateOption> GetBusinessExportTemplates(string projectId);

    Task<BusinessExportWorkbook> ExportBusinessWorkbookAsync(
        string projectId,
        string templateId,
        CancellationToken cancellationToken);
}
```

State that connectors which do not implement the extension still support config-only initialization.

- [ ] **Step 3: Update manual checklist**

In `docs/vsto-manual-test-checklist.md`, add these cases under Ribbon Sync validation:

```markdown
### 初始化当前表：业务模板导入

- 新建空白 workbook，选择项目并确认布局，点击 `初始化当前表`，默认选中 `从模板创建作业表`
- 确认后当前 sheet 保留原 sheet 名，内容来自导出 workbook 的 `Business Data`，活动单元格为 `A1`
- 查看 `xISDP_Setting`，当前 sheet 的 `SheetBindings + SheetFieldMappings` 已写入
- 在非空 sheet 点击 `初始化当前表`，默认选中 `仅初始化配置`
- 在非空 sheet 手动选择 `从模板创建作业表`，确认能看到覆盖风险提示和 `覆盖并初始化` 按钮
- 模板下载阶段点击取消，当前 sheet 内容不变，`xISDP_Setting` 不新增或改写该 sheet 的配置
- 切换到 `xISDP_Setting` 或 `xISDP_Log` 后点击 `初始化当前表`，看到禁止初始化提示
```

- [ ] **Step 4: Update module index**

In `docs/module-index.md`, add the new spec and plan under Ribbon Sync related design / plans:

```markdown
[docs/superpowers/specs/2026-06-25-initialize-sheet-business-template-import-design.md](./superpowers/specs/2026-06-25-initialize-sheet-business-template-import-design.md)<br>
[docs/superpowers/plans/2026-06-25-initialize-sheet-business-template-import-implementation-plan.md](./superpowers/plans/2026-06-25-initialize-sheet-business-template-import-implementation-plan.md)<br>
```

- [ ] **Step 5: Run all relevant automated tests**

Run:

```powershell
dotnet test tests/OfficeAgent.Core.Tests/OfficeAgent.Core.Tests.csproj
dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj
dotnet test tests/OfficeAgent.IntegrationTests/OfficeAgent.IntegrationTests.csproj
```

Expected: PASS.

- [ ] **Step 6: Run add-in build validation**

Run:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File eng/Dev-RefreshExcelAddIn.ps1 -CloseExcel
```

Expected: frontend build, Debug VSTO build, and local Excel registration refresh all pass.

- [ ] **Step 7: Manual validation**

Run:

```powershell
cd tests/mock-server
npm start
```

Then perform the checklist entries added in Step 3. Capture one screenshot of the initialization dialog on a blank sheet and one screenshot of the overwrite warning on a nonblank sheet for the PR description.

- [ ] **Step 8: Commit docs and validation checklist**

Run:

```powershell
git add docs/modules/ribbon-sync-current-behavior.md docs/ribbon-sync-real-system-integration-guide.md docs/vsto-manual-test-checklist.md docs/module-index.md
git commit -m "docs: document template-based initialization"
```

---

## Self-Review

Spec coverage:

- Blank sheet default template import is covered by Task 4 dialog defaults and Task 6 controller flow.
- Nonblank sheet default config-only plus overwrite warning is covered by Task 4 dialog state and localized risk text.
- Optional connector extension is covered by Task 1 and Task 2.
- `.xlsx` binary export and `Business Data` source sheet are covered by Task 2, Task 3, and Task 7.
- Metadata is not written before download/import succeeds because Task 1 introduces `PrepareSheetInitialization` and `SaveSheetInitialization`, and Task 5 writes after `ImportBusinessDataSheet`.
- Cancel during download is covered by Task 4 progress dialog and Task 6 controller cancellation path.
- Managed sheets are blocked by Task 6.
- Docs and manual checklist are covered by Task 8.

Placeholder scan:

- Every new type, method signature, endpoint, and command used by later tasks is defined earlier in the plan.
- The plan uses exact endpoint names, model property names, test names, and commit commands.

Type consistency:

- Core model names are `BusinessExportTemplateOption`, `BusinessExportWorkbook`, and `SheetInitializationPlan`.
- Optional connector interface name is `IBusinessExportTemplateConnector`.
- Dialog flow uses `InitializeSheetDialogRequest`, `InitializeSheetTemplateLoadResult`, `InitializeSheetDialogResult`, and `InitializeSheetMode`.
- Progress flow uses `IInitializeSheetImportProgress`.
- Excel copy boundary uses `IBusinessWorkbookImporter`.
