# Ribbon Sync Configurability Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Implement configurable Ribbon Sync sheet metadata, field-mapping initialization, and header-text-based upload/download flows while removing incremental upload from the initial release.

**Architecture:** Treat `_OfficeAgentMetadata` as the runtime source of truth for each managed sheet. Keep current-system-specific column names and seed generation inside the connector, but resolve live Excel columns from the current sheet headers on every upload/download so the execution path no longer depends on persisted column indexes or snapshot tables.

**Tech Stack:** C#, .NET Framework 4.8, VSTO, Excel Interop, WinForms, Newtonsoft.Json, xUnit, Node.js mock server

---

## File Structure

- `src/OfficeAgent.Core/Models/SheetBinding.cs`
  Responsibility: store one row of project binding plus `HeaderStartRow`, `HeaderRowCount`, and `DataStartRow`
- `src/OfficeAgent.Core/Models/FieldMappingSemanticRole.cs`
  Responsibility: enumerate internal semantic roles for dynamic mapping columns
- `src/OfficeAgent.Core/Models/FieldMappingColumnDefinition.cs`
  Responsibility: describe one connector-defined mapping-table column and the semantic role it carries
- `src/OfficeAgent.Core/Models/FieldMappingTableDefinition.cs`
  Responsibility: describe the dynamic shape of `SheetFieldMappings` for one `systemKey`
- `src/OfficeAgent.Core/Models/SheetFieldMappingRow.cs`
  Responsibility: store one metadata row of mapping values using dynamic column names
- `src/OfficeAgent.Core/Models/WorksheetRuntimeColumn.cs`
  Responsibility: represent one live Excel column after header-text recognition, including actual `ColumnIndex`
- `src/OfficeAgent.Core/Services/IWorksheetMetadataStore.cs`
  Responsibility: load and save `SheetBindings` and `SheetFieldMappings`
- `src/OfficeAgent.Core/Services/ISystemConnector.cs`
  Responsibility: expose connector seed defaults, mapping definitions, download, and upload
- `src/OfficeAgent.Core/Sync/FieldMappingValueAccessor.cs`
  Responsibility: translate semantic roles to dynamic mapping-table values
- `src/OfficeAgent.Core/Sync/WorksheetSyncService.cs`
  Responsibility: orchestrate initialization seed generation, download, and upload without snapshots
- `src/OfficeAgent.Infrastructure/Http/CurrentBusinessFieldMappingSeedBuilder.cs`
  Responsibility: turn current-system `/head` plus sample rows into mapping-table seed rows
- `src/OfficeAgent.Infrastructure/Http/CurrentBusinessSystemConnector.cs`
  Responsibility: provide seed defaults, dynamic mapping definition, seed rows, `/find`, and `/batchSave`
- `src/OfficeAgent.ExcelAddIn/Excel/WorksheetMetadataStore.cs`
  Responsibility: persist visible `SheetBindings` and `SheetFieldMappings` tables
- `src/OfficeAgent.ExcelAddIn/Excel/IWorksheetGridAdapter.cs`
  Responsibility: read/write worksheet cells plus range clearing and column scanning helpers
- `src/OfficeAgent.ExcelAddIn/Excel/ExcelWorksheetGridAdapter.cs`
  Responsibility: VSTO Excel implementation of the expanded grid adapter
- `src/OfficeAgent.ExcelAddIn/Excel/WorksheetSchemaLayoutService.cs`
  Responsibility: build header plans using configurable header rows
- `src/OfficeAgent.ExcelAddIn/Excel/WorksheetHeaderMatcher.cs`
  Responsibility: resolve current Excel columns from header text and mapping rows
- `src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs`
  Responsibility: coordinate initialization, full/partial download, and full/partial upload on a sheet
- `src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs`
  Responsibility: manage project selection, auto-initialize attempts, explicit initialize action, and the four remaining sync actions
- `src/OfficeAgent.ExcelAddIn/AgentRibbon.cs`
  Responsibility: wire the new initialize button and remove the incremental upload handler
- `src/OfficeAgent.ExcelAddIn/AgentRibbon.Designer.cs`
  Responsibility: place the initialize button under the project dropdown and remove the incremental upload button
- `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs`
  Responsibility: compose the updated services and refresh Ribbon state when the active sheet changes
- `tests/OfficeAgent.Infrastructure.Tests/CurrentBusinessSystemConnectorTests.cs`
  Responsibility: verify connector defaults, mapping definition, and batch-save behavior
- `tests/OfficeAgent.Infrastructure.Tests/CurrentBusinessFieldMappingSeedBuilderTests.cs`
  Responsibility: verify current-system seed rows for single and activity columns
- `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetMetadataStoreTests.cs`
  Responsibility: verify `SheetBindings` and `SheetFieldMappings` persistence
- `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetHeaderMatcherTests.cs`
  Responsibility: verify single-row and two-row header recognition
- `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSchemaLayoutServiceTests.cs`
  Responsibility: verify configurable header rendering
- `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs`
  Responsibility: verify initialization, configurable full download, and header-text-based upload/download
- `tests/OfficeAgent.ExcelAddIn.Tests/RibbonSyncControllerTests.cs`
  Responsibility: verify project selection, auto-try initialize, and explicit initialize behavior
- `tests/OfficeAgent.IntegrationTests/CurrentBusinessSystemConnectorIntegrationTests.cs`
  Responsibility: verify mock-backed seed generation and upload/download roundtrip
- `tests/mock-server/server.js`
  Responsibility: support `/head`, `/find`, and `/batchSave` for the new initialization flow
- `docs/modules/ribbon-sync-current-behavior.md`
  Responsibility: update the module snapshot after implementation
- `docs/module-index.md`
  Responsibility: point future sessions to the new design and implementation plan
- `docs/vsto-manual-test-checklist.md`
  Responsibility: document manual verification for initialize/download/upload behavior

### Task 1: Add Configurable Metadata and Connector Contracts

**Files:**
- Modify: `src/OfficeAgent.Core/Models/SheetBinding.cs`
- Create: `src/OfficeAgent.Core/Models/FieldMappingSemanticRole.cs`
- Create: `src/OfficeAgent.Core/Models/FieldMappingColumnDefinition.cs`
- Create: `src/OfficeAgent.Core/Models/FieldMappingTableDefinition.cs`
- Create: `src/OfficeAgent.Core/Models/SheetFieldMappingRow.cs`
- Modify: `src/OfficeAgent.Core/Services/IWorksheetMetadataStore.cs`
- Modify: `src/OfficeAgent.Core/Services/ISystemConnector.cs`
- Test: `tests/OfficeAgent.Infrastructure.Tests/CurrentBusinessSystemConnectorTests.cs`

- [ ] **Step 1: Write the failing contract tests**

```csharp
[Fact]
public void CreateBindingSeedUsesConfigurableDefaults()
{
    var connector = new CurrentBusinessSystemConnector(
        () => new AppSettings { BusinessBaseUrl = "https://business.internal.example" },
        new HttpClient(new RecordingHandler()));

    var project = connector.GetProjects().Single();
    var binding = connector.CreateBindingSeed("Sheet1", project);

    Assert.Equal("Sheet1", binding.SheetName);
    Assert.Equal(1, binding.HeaderStartRow);
    Assert.Equal(2, binding.HeaderRowCount);
    Assert.Equal(3, binding.DataStartRow);
}

[Fact]
public void GetFieldMappingDefinitionExposesCurrentSystemSemanticRoles()
{
    var connector = new CurrentBusinessSystemConnector(
        () => new AppSettings { BusinessBaseUrl = "https://business.internal.example" },
        new HttpClient(new RecordingHandler()));

    var definition = connector.GetFieldMappingDefinition("performance");

    Assert.Contains(definition.Columns, column => column.ColumnName == "HeaderId" && column.Role == FieldMappingSemanticRole.HeaderIdentity);
    Assert.Contains(definition.Columns, column => column.ColumnName == "CurrentParentDisplayName" && column.Role == FieldMappingSemanticRole.CurrentParentHeaderText);
    Assert.Contains(definition.Columns, column => column.ColumnName == "PropertyId" && column.Role == FieldMappingSemanticRole.PropertyIdentity);
}
```

- [ ] **Step 2: Run the targeted infrastructure tests to confirm the interface gap**

Run: `dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj --filter FullyQualifiedName~CurrentBusinessSystemConnectorTests.CreateBindingSeedUsesConfigurableDefaults|FullyQualifiedName~CurrentBusinessSystemConnectorTests.GetFieldMappingDefinitionExposesCurrentSystemSemanticRoles`

Expected: FAIL with compiler errors for `CreateBindingSeed`, `GetFieldMappingDefinition`, `FieldMappingSemanticRole`, or the expanded `SheetBinding`.

- [ ] **Step 3: Add the new models and interface surface**

```csharp
namespace OfficeAgent.Core.Models
{
    public sealed class SheetBinding
    {
        public string SheetName { get; set; } = string.Empty;
        public string SystemKey { get; set; } = string.Empty;
        public string ProjectId { get; set; } = string.Empty;
        public string ProjectName { get; set; } = string.Empty;
        public int HeaderStartRow { get; set; } = 1;
        public int HeaderRowCount { get; set; } = 2;
        public int DataStartRow { get; set; } = 3;
    }

    public enum FieldMappingSemanticRole
    {
        HeaderIdentity,
        HeaderType,
        ApiFieldKey,
        IsIdColumn,
        DefaultSingleHeaderText,
        CurrentSingleHeaderText,
        DefaultParentHeaderText,
        CurrentParentHeaderText,
        DefaultChildHeaderText,
        CurrentChildHeaderText,
        ActivityIdentity,
        PropertyIdentity,
        AuxiliaryIdentity,
    }

    public sealed class FieldMappingColumnDefinition
    {
        public string ColumnName { get; set; } = string.Empty;
        public FieldMappingSemanticRole Role { get; set; }
        public string RoleKey { get; set; } = string.Empty;
    }

    public sealed class FieldMappingTableDefinition
    {
        public string SystemKey { get; set; } = string.Empty;
        public FieldMappingColumnDefinition[] Columns { get; set; } = Array.Empty<FieldMappingColumnDefinition>();
    }

    public sealed class SheetFieldMappingRow
    {
        public string SheetName { get; set; } = string.Empty;
        public IReadOnlyDictionary<string, string> Values { get; set; } =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    }
}
```

```csharp
public interface IWorksheetMetadataStore
{
    void SaveBinding(SheetBinding binding);
    SheetBinding LoadBinding(string sheetName);
    void SaveFieldMappings(string sheetName, FieldMappingTableDefinition definition, IReadOnlyList<SheetFieldMappingRow> rows);
    SheetFieldMappingRow[] LoadFieldMappings(string sheetName, FieldMappingTableDefinition definition);
    void ClearFieldMappings(string sheetName);
}
```

```csharp
public interface ISystemConnector
{
    IReadOnlyList<ProjectOption> GetProjects();
    SheetBinding CreateBindingSeed(string sheetName, ProjectOption project);
    FieldMappingTableDefinition GetFieldMappingDefinition(string projectId);
    IReadOnlyList<SheetFieldMappingRow> BuildFieldMappingSeed(string sheetName, string projectId);
    IReadOnlyList<IDictionary<string, object>> Find(string projectId, IReadOnlyList<string> rowIds, IReadOnlyList<string> fieldKeys);
    void BatchSave(string projectId, IReadOnlyList<CellChange> changes);
}
```

- [ ] **Step 4: Run the targeted tests and the full infrastructure suite**

Run: `dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj --filter FullyQualifiedName~CurrentBusinessSystemConnectorTests`

Expected: PASS for the two new contract tests and the existing connector tests after the stubs are in place.

Run: `dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj`

Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/OfficeAgent.Core/Models/SheetBinding.cs src/OfficeAgent.Core/Models/FieldMappingSemanticRole.cs src/OfficeAgent.Core/Models/FieldMappingColumnDefinition.cs src/OfficeAgent.Core/Models/FieldMappingTableDefinition.cs src/OfficeAgent.Core/Models/SheetFieldMappingRow.cs src/OfficeAgent.Core/Services/IWorksheetMetadataStore.cs src/OfficeAgent.Core/Services/ISystemConnector.cs tests/OfficeAgent.Infrastructure.Tests/CurrentBusinessSystemConnectorTests.cs
git commit -m "feat: add configurable ribbon sync contracts"
```

### Task 2: Rewrite Visible Metadata Storage Around `SheetBindings` and `SheetFieldMappings`

**Files:**
- Modify: `src/OfficeAgent.ExcelAddIn/Excel/WorksheetMetadataStore.cs`
- Modify: `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetMetadataStoreTests.cs`

- [ ] **Step 1: Write the failing metadata-store tests**

```csharp
[Fact]
public void SaveBindingRoundTripsLayoutConfiguration()
{
    var (store, adapter) = CreateStore();
    var binding = new SheetBinding
    {
        SheetName = "Sheet1",
        SystemKey = "current-business-system",
        ProjectId = "performance",
        ProjectName = "绩效项目",
        HeaderStartRow = 3,
        HeaderRowCount = 2,
        DataStartRow = 6,
    };

    InvokeSaveBinding(store, binding);
    var loaded = InvokeLoadBinding(store, "Sheet1");

    Assert.Equal(3, loaded.HeaderStartRow);
    Assert.Equal(2, loaded.HeaderRowCount);
    Assert.Equal(6, loaded.DataStartRow);
}

[Fact]
public void SaveFieldMappingsPreservesOtherSheetsAndUsesDynamicHeaders()
{
    var (store, adapter) = CreateStore();
    var definition = new FieldMappingTableDefinition
    {
        SystemKey = "current-business-system",
        Columns = new[]
        {
            new FieldMappingColumnDefinition { ColumnName = "HeaderId", Role = FieldMappingSemanticRole.HeaderIdentity },
            new FieldMappingColumnDefinition { ColumnName = "CurrentSingleDisplayName", Role = FieldMappingSemanticRole.CurrentSingleHeaderText },
        },
    };

    adapter.SeedTable("SheetFieldMappings", new[]
    {
        new[] { "SheetA", "legacy_id", "旧列名" },
    });

    InvokeSaveFieldMappings(
        store,
        "Sheet1",
        definition,
        new[]
        {
            new SheetFieldMappingRow
            {
                SheetName = "Sheet1",
                Values = new Dictionary<string, string>
                {
                    ["HeaderId"] = "owner_name",
                    ["CurrentSingleDisplayName"] = "项目负责人",
                },
            },
        });

    var loaded = InvokeLoadFieldMappings(store, "Sheet1", definition);
    Assert.Single(loaded);
    Assert.Equal("owner_name", loaded[0].Values["HeaderId"]);

    var rawRows = adapter.ReadSeededTable("SheetFieldMappings");
    Assert.Contains(rawRows, row => row[0] == "SheetA" && row[1] == "legacy_id");
}
```

- [ ] **Step 2: Run the metadata-store tests to confirm the old snapshot-based implementation no longer fits**

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~WorksheetMetadataStoreTests.SaveBindingRoundTripsLayoutConfiguration|FullyQualifiedName~WorksheetMetadataStoreTests.SaveFieldMappingsPreservesOtherSheetsAndUsesDynamicHeaders`

Expected: FAIL with missing methods or row-length assertions against the old `SheetBindings` and `SheetSnapshots` logic.

- [ ] **Step 3: Implement the new metadata tables**

```csharp
private static readonly string[] BindingHeaders =
{
    "SheetName",
    "SystemKey",
    "ProjectId",
    "ProjectName",
    "HeaderStartRow",
    "HeaderRowCount",
    "DataStartRow",
};

public void SaveBinding(SheetBinding binding)
{
    adapter.EnsureWorksheet(MetadataSheetName, visible: true);
    var rows = adapter.ReadTable(BindingsTableName)?.ToList() ?? new List<string[]>();
    rows.RemoveAll(row => row.Length > 0 && string.Equals(row[0], binding.SheetName, StringComparison.OrdinalIgnoreCase));
    rows.Add(new[]
    {
        binding.SheetName ?? string.Empty,
        binding.SystemKey ?? string.Empty,
        binding.ProjectId ?? string.Empty,
        binding.ProjectName ?? string.Empty,
        binding.HeaderStartRow.ToString(),
        binding.HeaderRowCount.ToString(),
        binding.DataStartRow.ToString(),
    });
    adapter.WriteTable(BindingsTableName, BindingHeaders, rows.ToArray());
}
```

```csharp
public void SaveFieldMappings(string sheetName, FieldMappingTableDefinition definition, IReadOnlyList<SheetFieldMappingRow> rows)
{
    adapter.EnsureWorksheet(MetadataSheetName, visible: true);
    var headers = new[] { "SheetName" }
        .Concat((definition?.Columns ?? Array.Empty<FieldMappingColumnDefinition>()).Select(column => column.ColumnName))
        .ToArray();

    var existing = adapter.ReadTable(FieldMappingsTableName)?.ToList() ?? new List<string[]>();
    existing.RemoveAll(row => row.Length > 0 && string.Equals(row[0], sheetName, StringComparison.OrdinalIgnoreCase));

    foreach (var row in rows ?? Array.Empty<SheetFieldMappingRow>())
    {
        var values = new List<string> { sheetName };
        foreach (var column in definition.Columns)
        {
            values.Add(row.Values.TryGetValue(column.ColumnName, out var value) ? value : string.Empty);
        }

        existing.Add(values.ToArray());
    }

    adapter.WriteTable(FieldMappingsTableName, headers, existing.ToArray());
}

public SheetFieldMappingRow[] LoadFieldMappings(string sheetName, FieldMappingTableDefinition definition)
{
    var rows = adapter.ReadTable(FieldMappingsTableName) ?? Array.Empty<string[]>();
    return rows
        .Where(row => row.Length > 0 && string.Equals(row[0], sheetName, StringComparison.OrdinalIgnoreCase))
        .Select(row =>
        {
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            for (var index = 0; index < definition.Columns.Length; index++)
            {
                values[definition.Columns[index].ColumnName] = row.Length > index + 1 ? row[index + 1] : string.Empty;
            }

            return new SheetFieldMappingRow
            {
                SheetName = sheetName,
                Values = values,
            };
        })
        .ToArray();
}
```

- [ ] **Step 4: Run the add-in test suite**

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~WorksheetMetadataStoreTests`

Expected: PASS

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj`

Expected: PASS, with the snapshot-specific tests removed or rewritten.

- [ ] **Step 5: Commit**

```bash
git add src/OfficeAgent.ExcelAddIn/Excel/WorksheetMetadataStore.cs tests/OfficeAgent.ExcelAddIn.Tests/WorksheetMetadataStoreTests.cs
git commit -m "feat: persist configurable ribbon sync metadata tables"
```

### Task 3: Build Current-System Field-Mapping Definitions and Seed Rows

**Files:**
- Create: `src/OfficeAgent.Infrastructure/Http/CurrentBusinessFieldMappingSeedBuilder.cs`
- Modify: `src/OfficeAgent.Infrastructure/Http/CurrentBusinessSystemConnector.cs`
- Create: `tests/OfficeAgent.Infrastructure.Tests/CurrentBusinessFieldMappingSeedBuilderTests.cs`
- Modify: `tests/OfficeAgent.Infrastructure.Tests/CurrentBusinessSystemConnectorTests.cs`

- [ ] **Step 1: Write the failing seed-builder tests**

```csharp
[Fact]
public void BuildCreatesSingleAndActivityRowsUsingCurrentSystemColumnNames()
{
    var builder = new CurrentBusinessFieldMappingSeedBuilder();
    var rows = builder.Build(
        "Sheet1",
        new[]
        {
            new CurrentBusinessHeadDefinition { FieldKey = "row_id", HeaderText = "ID", IsId = true },
            new CurrentBusinessHeadDefinition { FieldKey = "owner_name", HeaderText = "负责人" },
            new CurrentBusinessHeadDefinition { HeadType = "activity", ActivityId = "12345678", ActivityName = "测试活动111" },
        },
        new[]
        {
            new Dictionary<string, object>
            {
                ["row_id"] = "row-1",
                ["owner_name"] = "张三",
                ["start_12345678"] = "2026-01-02",
                ["end_12345678"] = "2026-01-05",
            },
        });

    Assert.Contains(rows, row => row.Values["HeaderId"] == "row_id" && row.Values["IsIdColumn"] == "true");
    Assert.Contains(rows, row => row.Values["ApiFieldKey"] == "start_12345678" && row.Values["CurrentParentDisplayName"] == "测试活动111");
    Assert.Contains(rows, row => row.Values["ApiFieldKey"] == "end_12345678" && row.Values["CurrentChildDisplayName"] == "结束时间");
}
```

```csharp
[Fact]
public void BuildFieldMappingSeedCallsHeadAndFindBeforeReturningRows()
{
    var handler = new SeedAwareHandler();
    var connector = CurrentBusinessSystemConnector.ForTests("https://api.internal.example", handler);

    var rows = connector.BuildFieldMappingSeed("Sheet1", "performance");

    Assert.NotEmpty(rows);
    Assert.Equal(new[] { "/head", "/find" }, handler.Paths);
}
```

- [ ] **Step 2: Run the targeted infrastructure tests**

Run: `dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj --filter FullyQualifiedName~CurrentBusinessFieldMappingSeedBuilderTests|FullyQualifiedName~CurrentBusinessSystemConnectorTests.BuildFieldMappingSeedCallsHeadAndFindBeforeReturningRows`

Expected: FAIL with missing `CurrentBusinessFieldMappingSeedBuilder` or `BuildFieldMappingSeed`.

- [ ] **Step 3: Implement the seed builder and connector methods**

```csharp
public sealed class CurrentBusinessFieldMappingSeedBuilder
{
    private static readonly IReadOnlyDictionary<string, string> PropertyLabels =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["start"] = "开始时间",
            ["end"] = "结束时间",
        };

    public IReadOnlyList<SheetFieldMappingRow> Build(
        string sheetName,
        IReadOnlyList<CurrentBusinessHeadDefinition> headList,
        IReadOnlyList<IDictionary<string, object>> rows)
    {
        var result = new List<SheetFieldMappingRow>();

        foreach (var head in headList.Where(item => !string.Equals(item.HeadType, "activity", StringComparison.OrdinalIgnoreCase)))
        {
            result.Add(CreateRow(
                sheetName,
                headerId: head.FieldKey,
                headerType: "single",
                apiFieldKey: head.FieldKey,
                isIdColumn: head.IsId,
                currentSingle: head.HeaderText));
        }

        var activityHeads = headList
            .Where(item => string.Equals(item.HeadType, "activity", StringComparison.OrdinalIgnoreCase))
            .ToDictionary(item => item.ActivityId, StringComparer.OrdinalIgnoreCase);

        var flatKeys = rows.SelectMany(row => row.Keys).Distinct(StringComparer.OrdinalIgnoreCase);
        foreach (var flatKey in flatKeys.Where(key => key.Contains("_")))
        {
            var parts = flatKey.Split(new[] { '_' }, 2);
            if (parts.Length != 2 || !activityHeads.TryGetValue(parts[1], out var activity))
            {
                continue;
            }

            result.Add(CreateRow(
                sheetName,
                headerId: flatKey,
                headerType: "activityProperty",
                apiFieldKey: flatKey,
                isIdColumn: false,
                currentParent: activity.ActivityName,
                currentChild: PropertyLabels.TryGetValue(parts[0], out var label) ? label : parts[0],
                activityId: activity.ActivityId,
                propertyId: parts[0]));
        }

        return result;
    }
}
```

```csharp
public FieldMappingTableDefinition GetFieldMappingDefinition(string projectId)
{
    return new FieldMappingTableDefinition
    {
        SystemKey = "current-business-system",
        Columns = new[]
        {
            new FieldMappingColumnDefinition { ColumnName = "HeaderId", Role = FieldMappingSemanticRole.HeaderIdentity },
            new FieldMappingColumnDefinition { ColumnName = "HeaderType", Role = FieldMappingSemanticRole.HeaderType },
            new FieldMappingColumnDefinition { ColumnName = "ApiFieldKey", Role = FieldMappingSemanticRole.ApiFieldKey },
            new FieldMappingColumnDefinition { ColumnName = "IsIdColumn", Role = FieldMappingSemanticRole.IsIdColumn },
            new FieldMappingColumnDefinition { ColumnName = "DefaultSingleDisplayName", Role = FieldMappingSemanticRole.DefaultSingleHeaderText },
            new FieldMappingColumnDefinition { ColumnName = "CurrentSingleDisplayName", Role = FieldMappingSemanticRole.CurrentSingleHeaderText },
            new FieldMappingColumnDefinition { ColumnName = "DefaultParentDisplayName", Role = FieldMappingSemanticRole.DefaultParentHeaderText },
            new FieldMappingColumnDefinition { ColumnName = "CurrentParentDisplayName", Role = FieldMappingSemanticRole.CurrentParentHeaderText },
            new FieldMappingColumnDefinition { ColumnName = "DefaultChildDisplayName", Role = FieldMappingSemanticRole.DefaultChildHeaderText },
            new FieldMappingColumnDefinition { ColumnName = "CurrentChildDisplayName", Role = FieldMappingSemanticRole.CurrentChildHeaderText },
            new FieldMappingColumnDefinition { ColumnName = "ActivityId", Role = FieldMappingSemanticRole.ActivityIdentity },
            new FieldMappingColumnDefinition { ColumnName = "PropertyId", Role = FieldMappingSemanticRole.PropertyIdentity },
        },
    };
}
```

```csharp
public IReadOnlyList<SheetFieldMappingRow> BuildFieldMappingSeed(string sheetName, string projectId)
{
    var headWrapper = Post<SchemaHeadWrapper>("/head", new { projectId });
    var headList = headWrapper?.HeadList ?? Array.Empty<CurrentBusinessHeadDefinition>();
    var sampleRows = Find(projectId, Array.Empty<string>(), Array.Empty<string>());
    return fieldMappingSeedBuilder.Build(sheetName, headList, sampleRows);
}

public SheetBinding CreateBindingSeed(string sheetName, ProjectOption project)
{
    return new SheetBinding
    {
        SheetName = sheetName ?? string.Empty,
        SystemKey = project?.SystemKey ?? string.Empty,
        ProjectId = project?.ProjectId ?? string.Empty,
        ProjectName = project?.DisplayName ?? string.Empty,
        HeaderStartRow = 1,
        HeaderRowCount = 2,
        DataStartRow = 3,
    };
}
```

- [ ] **Step 4: Run the infrastructure suite**

Run: `dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj --filter FullyQualifiedName~CurrentBusinessFieldMappingSeedBuilderTests`

Expected: PASS

Run: `dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj`

Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/OfficeAgent.Infrastructure/Http/CurrentBusinessFieldMappingSeedBuilder.cs src/OfficeAgent.Infrastructure/Http/CurrentBusinessSystemConnector.cs tests/OfficeAgent.Infrastructure.Tests/CurrentBusinessFieldMappingSeedBuilderTests.cs tests/OfficeAgent.Infrastructure.Tests/CurrentBusinessSystemConnectorTests.cs
git commit -m "feat: add current business field mapping seeds"
```

### Task 4: Add Configurable Header Layout and Live Header Matching

**Files:**
- Create: `src/OfficeAgent.Core/Models/WorksheetRuntimeColumn.cs`
- Create: `src/OfficeAgent.Core/Sync/FieldMappingValueAccessor.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/Excel/IWorksheetGridAdapter.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/Excel/ExcelWorksheetGridAdapter.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/Excel/WorksheetSchemaLayoutService.cs`
- Create: `src/OfficeAgent.ExcelAddIn/Excel/WorksheetHeaderMatcher.cs`
- Modify: `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSchemaLayoutServiceTests.cs`
- Create: `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetHeaderMatcherTests.cs`

- [ ] **Step 1: Write the failing layout and matcher tests**

```csharp
[Fact]
public void BuildHeaderPlanHonorsHeaderStartRowAndSingleRowLayout()
{
    var service = CreateService();
    var binding = new SheetBinding { HeaderStartRow = 3, HeaderRowCount = 1 };
    var columns = new[]
    {
        new WorksheetRuntimeColumn { ColumnIndex = 1, ApiFieldKey = "row_id", HeaderType = "single", DisplayText = "ID", IsIdColumn = true },
        new WorksheetRuntimeColumn { ColumnIndex = 2, ApiFieldKey = "owner_name", HeaderType = "single", DisplayText = "项目负责人" },
    };

    var plan = BuildHeaderPlan(service, binding, columns);

    Assert.Contains(plan, cell => cell.Row == 3 && cell.Column == 1 && cell.RowSpan == 1 && cell.Text == "ID");
    Assert.DoesNotContain(plan, cell => cell.Row == 4);
}
```

```csharp
[Fact]
public void MatchUsesParentAndChildDisplayNamesForTwoRowHeaders()
{
    var grid = new FakeGrid();
    grid.SetCell("Sheet1", 3, 1, "ID");
    grid.SetCell("Sheet1", 3, 2, "测试活动111");
    grid.SetCell("Sheet1", 4, 2, "开始时间");
    grid.SetCell("Sheet1", 3, 3, "测试活动111");
    grid.SetCell("Sheet1", 4, 3, "结束时间");

    var matcher = new WorksheetHeaderMatcher(new FieldMappingValueAccessor());
    var binding = new SheetBinding { SheetName = "Sheet1", HeaderStartRow = 3, HeaderRowCount = 2 };
    var definition = BuildDefinition();
    var mappings = BuildActivityMappings("Sheet1");

    var columns = matcher.Match("Sheet1", binding, definition, mappings, grid);

    Assert.Contains(columns, column => column.ColumnIndex == 2 && column.ApiFieldKey == "start_12345678");
    Assert.Contains(columns, column => column.ColumnIndex == 3 && column.ApiFieldKey == "end_12345678");
}
```

- [ ] **Step 2: Run the layout-related add-in tests**

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~WorksheetSchemaLayoutServiceTests|FullyQualifiedName~WorksheetHeaderMatcherTests`

Expected: FAIL because the grid adapter lacks column scanning and the current layout service assumes row `1/2`.

- [ ] **Step 3: Implement runtime column matching and configurable header layout**

```csharp
public sealed class WorksheetRuntimeColumn
{
    public int ColumnIndex { get; set; }
    public string ApiFieldKey { get; set; } = string.Empty;
    public string HeaderType { get; set; } = string.Empty;
    public string DisplayText { get; set; } = string.Empty;
    public string ParentDisplayText { get; set; } = string.Empty;
    public string ChildDisplayText { get; set; } = string.Empty;
    public bool IsIdColumn { get; set; }
}
```

```csharp
internal interface IWorksheetGridAdapter
{
    string GetCellText(string sheetName, int row, int column);
    void SetCellText(string sheetName, int row, int column, string value);
    void ClearRange(string sheetName, int startRow, int endRow, int startColumn, int endColumn);
    void MergeCells(string sheetName, int row, int column, int rowSpan, int columnSpan);
    int GetLastUsedRow(string sheetName);
    int GetLastUsedColumn(string sheetName);
}
```

```csharp
public HeaderCellPlan[] BuildHeaderPlan(SheetBinding binding, IReadOnlyList<WorksheetRuntimeColumn> columns)
{
    var startRow = binding.HeaderStartRow;
    if (binding.HeaderRowCount == 1)
    {
        return columns.Select(column => new HeaderCellPlan
        {
            Row = startRow,
            Column = column.ColumnIndex,
            RowSpan = 1,
            ColumnSpan = 1,
            Text = column.DisplayText,
        }).ToArray();
    }

    var cells = new List<HeaderCellPlan>();
    foreach (var column in columns.Where(item => string.Equals(item.HeaderType, "single", StringComparison.OrdinalIgnoreCase)))
    {
        cells.Add(new HeaderCellPlan
        {
            Row = startRow,
            Column = column.ColumnIndex,
            RowSpan = 2,
            Text = column.DisplayText,
        });
    }

    foreach (var group in columns.Where(item => string.Equals(item.HeaderType, "activityProperty", StringComparison.OrdinalIgnoreCase)).GroupBy(item => item.ParentDisplayText))
    {
        var ordered = group.OrderBy(item => item.ColumnIndex).ToArray();
        cells.Add(new HeaderCellPlan
        {
            Row = startRow,
            Column = ordered[0].ColumnIndex,
            ColumnSpan = ordered.Length,
            Text = ordered[0].ParentDisplayText,
        });

        foreach (var column in ordered)
        {
            cells.Add(new HeaderCellPlan
            {
                Row = startRow + 1,
                Column = column.ColumnIndex,
                Text = column.ChildDisplayText,
            });
        }
    }

    return cells.ToArray();
}
```

```csharp
public WorksheetRuntimeColumn[] Match(
    string sheetName,
    SheetBinding binding,
    FieldMappingTableDefinition definition,
    IReadOnlyList<SheetFieldMappingRow> mappings,
    IWorksheetGridAdapter grid)
{
    var result = new List<WorksheetRuntimeColumn>();
    var lastUsedColumn = grid.GetLastUsedColumn(sheetName);
    var currentParent = string.Empty;

    for (var column = 1; column <= lastUsedColumn; column++)
    {
        var topText = grid.GetCellText(sheetName, binding.HeaderStartRow, column);
        var bottomText = binding.HeaderRowCount > 1
            ? grid.GetCellText(sheetName, binding.HeaderStartRow + 1, column)
            : string.Empty;

        if (!string.IsNullOrWhiteSpace(topText))
        {
            currentParent = topText;
        }

        var match = FindMatch(definition, mappings, topText, bottomText, currentParent, binding.HeaderRowCount);
        if (match == null)
        {
            continue;
        }

        result.Add(match);
    }

    return result.ToArray();
}
```

- [ ] **Step 4: Run the targeted add-in tests**

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~WorksheetSchemaLayoutServiceTests|FullyQualifiedName~WorksheetHeaderMatcherTests`

Expected: PASS

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~WorksheetSelectionResolverTests`

Expected: PASS, because the resolver still works once it receives live `ColumnIndex` values from the matcher.

- [ ] **Step 5: Commit**

```bash
git add src/OfficeAgent.Core/Models/WorksheetRuntimeColumn.cs src/OfficeAgent.Core/Sync/FieldMappingValueAccessor.cs src/OfficeAgent.ExcelAddIn/Excel/IWorksheetGridAdapter.cs src/OfficeAgent.ExcelAddIn/Excel/ExcelWorksheetGridAdapter.cs src/OfficeAgent.ExcelAddIn/Excel/WorksheetSchemaLayoutService.cs src/OfficeAgent.ExcelAddIn/Excel/WorksheetHeaderMatcher.cs tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSchemaLayoutServiceTests.cs tests/OfficeAgent.ExcelAddIn.Tests/WorksheetHeaderMatcherTests.cs
git commit -m "feat: match ribbon sync columns from live headers"
```

### Task 5: Rework Initialization, Full Download, and Upload Execution

**Files:**
- Modify: `src/OfficeAgent.Core/Sync/WorksheetSyncService.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs`
- Modify: `tests/OfficeAgent.Core.Tests/WorksheetSyncServiceTests.cs`
- Modify: `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs`

- [ ] **Step 1: Write the failing execution tests**

```csharp
[Fact]
public void InitializeCurrentSheetWritesBindingAndFieldMappingsWithoutTouchingBusinessCells()
{
    var connector = new FakeSystemConnector();
    var metadataStore = new FakeWorksheetMetadataStore();
    var selectionReader = new FakeWorksheetSelectionReader();
    var (service, grid) = CreateService(connector, metadataStore, selectionReader);

    grid.SetCell("Sheet1", 1, 1, "现有说明");

    InvokeInitialize(service, "Sheet1", new ProjectOption
    {
        SystemKey = "current-business-system",
        ProjectId = "performance",
        DisplayName = "绩效项目",
    });

    Assert.Equal("现有说明", grid.GetCell("Sheet1", 1, 1));
    Assert.Equal(1, metadataStore.LastSavedBinding.HeaderStartRow);
    Assert.NotEmpty(metadataStore.LastSavedFieldMappings);
}

[Fact]
public void ExecuteFullDownloadHonorsConfiguredHeaderAndDataRows()
{
    var connector = new FakeSystemConnector();
    var metadataStore = new FakeWorksheetMetadataStore();
    metadataStore.Bindings["Sheet1"] = new SheetBinding
    {
        SheetName = "Sheet1",
        SystemKey = "current-business-system",
        ProjectId = "performance",
        ProjectName = "绩效项目",
        HeaderStartRow = 3,
        HeaderRowCount = 2,
        DataStartRow = 6,
    };
    metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");
    connector.FindResult = new[] { CreateRow("row-1", "张三", "2026-01-02", "2026-01-05") };

    var (service, grid) = CreateService(connector, metadataStore, new FakeWorksheetSelectionReader());
    grid.SetCell("Sheet1", 1, 1, "统计说明");
    grid.SetCell("Sheet1", 5, 1, "统计行");

    var plan = InvokePrepare(service, "PrepareFullDownload", "Sheet1");
    InvokeExecute(service, "ExecuteDownload", plan);

    Assert.Equal("统计说明", grid.GetCell("Sheet1", 1, 1));
    Assert.Equal("统计行", grid.GetCell("Sheet1", 5, 1));
    Assert.Equal("ID", grid.GetCell("Sheet1", 3, 1));
    Assert.Equal("row-1", grid.GetCell("Sheet1", 6, 1));
}
```

- [ ] **Step 2: Run the targeted core and add-in tests**

Run: `dotnet test tests/OfficeAgent.Core.Tests/OfficeAgent.Core.Tests.csproj --filter FullyQualifiedName~WorksheetSyncServiceTests`

Expected: FAIL because the service still expects snapshots and schema lookup.

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~WorksheetSyncExecutionServiceTests.InitializeCurrentSheetWritesBindingAndFieldMappingsWithoutTouchingBusinessCells|FullyQualifiedName~WorksheetSyncExecutionServiceTests.ExecuteFullDownloadHonorsConfiguredHeaderAndDataRows`

Expected: FAIL because the execution service still uses `HeaderRowCount = 2`, `DataStartRow = 3`, and snapshot methods.

- [ ] **Step 3: Implement initialization and header-text-based execution**

```csharp
public sealed class WorksheetSyncService
{
    private readonly ISystemConnector connector;
    private readonly IWorksheetMetadataStore metadataStore;

    public void InitializeSheet(string sheetName, ProjectOption project)
    {
        var binding = connector.CreateBindingSeed(sheetName, project);
        var definition = connector.GetFieldMappingDefinition(project.ProjectId);
        var seedRows = connector.BuildFieldMappingSeed(sheetName, project.ProjectId);

        metadataStore.SaveBinding(binding);
        metadataStore.SaveFieldMappings(sheetName, definition, seedRows);
    }
    
    public SheetBinding LoadBinding(string sheetName)
    {
        return metadataStore.LoadBinding(sheetName);
    }

    public FieldMappingTableDefinition LoadFieldMappingDefinition(string projectId)
    {
        return connector.GetFieldMappingDefinition(projectId);
    }

    public SheetFieldMappingRow[] LoadFieldMappings(string sheetName, string projectId)
    {
        var definition = connector.GetFieldMappingDefinition(projectId);
        return metadataStore.LoadFieldMappings(sheetName, definition);
    }

    public IReadOnlyList<IDictionary<string, object>> Download(string projectId, IReadOnlyList<string> rowIds, IReadOnlyList<string> fieldKeys)
    {
        return connector.Find(projectId, rowIds, fieldKeys);
    }

    public void Upload(string projectId, IReadOnlyList<CellChange> changes)
    {
        connector.BatchSave(projectId, changes);
    }
}
```

```csharp
public void InitializeCurrentSheet(string sheetName, ProjectOption project)
{
    worksheetSyncService.InitializeSheet(sheetName, project);
}

private WorksheetRuntimeColumn[] LoadRuntimeColumns(string sheetName, SheetBinding binding)
{
    var definition = worksheetSyncService.LoadFieldMappingDefinition(binding.ProjectId);
    var mappings = worksheetSyncService.LoadFieldMappings(sheetName, binding.ProjectId);
    return headerMatcher.Match(sheetName, binding, definition, mappings, gridAdapter);
}

private void WriteFullWorksheet(WorksheetDownloadPlan plan)
{
    var binding = plan.Binding;
    var columns = plan.RuntimeColumns;
    var headerPlan = layoutService.BuildHeaderPlan(binding, columns);
    var lastColumn = columns.Length == 0 ? 0 : columns.Max(column => column.ColumnIndex);
    var clearEndRow = Math.Max(gridAdapter.GetLastUsedRow(plan.SheetName), binding.DataStartRow + plan.Rows.Count + 10);

    gridAdapter.ClearRange(plan.SheetName, binding.HeaderStartRow, binding.HeaderStartRow + binding.HeaderRowCount - 1, 1, lastColumn);
    gridAdapter.ClearRange(plan.SheetName, binding.DataStartRow, clearEndRow, 1, lastColumn);

    foreach (var headerCell in headerPlan)
    {
        gridAdapter.SetCellText(plan.SheetName, headerCell.Row, headerCell.Column, headerCell.Text);
        gridAdapter.MergeCells(plan.SheetName, headerCell.Row, headerCell.Column, headerCell.RowSpan, headerCell.ColumnSpan);
    }

    for (var rowIndex = 0; rowIndex < plan.Rows.Count; rowIndex++)
    {
        var targetRow = binding.DataStartRow + rowIndex;
        foreach (var column in columns)
        {
            gridAdapter.SetCellText(plan.SheetName, targetRow, column.ColumnIndex, GetRowValue(plan.Rows[rowIndex], column.ApiFieldKey));
        }
    }
}
```

- [ ] **Step 4: Run the rewritten execution tests**

Run: `dotnet test tests/OfficeAgent.Core.Tests/OfficeAgent.Core.Tests.csproj --filter FullyQualifiedName~WorksheetSyncServiceTests`

Expected: PASS after the snapshot-based tests are replaced with initialization and direct upload/download orchestration coverage.

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~WorksheetSyncExecutionServiceTests`

Expected: PASS after the incremental-upload tests are removed or rewritten for initialize/full/partial flows.

- [ ] **Step 5: Commit**

```bash
git add src/OfficeAgent.Core/Sync/WorksheetSyncService.cs src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs tests/OfficeAgent.Core.Tests/WorksheetSyncServiceTests.cs tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs
git commit -m "feat: add configurable ribbon sync execution flow"
```

### Task 6: Update Ribbon UI, Controller Flow, and Add-In Wiring

**Files:**
- Modify: `src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/AgentRibbon.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/AgentRibbon.Designer.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs`
- Modify: `tests/OfficeAgent.ExcelAddIn.Tests/RibbonSyncControllerTests.cs`

- [ ] **Step 1: Write the failing controller and Ribbon tests**

```csharp
[Fact]
public void SelectProjectSavesBindingAndAttemptsAutoInitialize()
{
    var executionService = new FakeWorksheetSyncExecutionService();
    var metadataStore = new FakeWorksheetMetadataStore();
    var controller = CreateController(executionService, metadataStore, activeSheetName: "Sheet1");

    controller.SelectProject(new ProjectOption
    {
        SystemKey = "current-business-system",
        ProjectId = "performance",
        DisplayName = "绩效项目",
    });

    Assert.Equal("performance", metadataStore.LastSavedBinding.ProjectId);
    Assert.Equal("Sheet1", executionService.LastAutoInitializeSheetName);
}

[Fact]
public void ExecuteInitializeCurrentSheetCallsExecutionService()
{
    var executionService = new FakeWorksheetSyncExecutionService();
    var metadataStore = new FakeWorksheetMetadataStore();
    metadataStore.Bindings["Sheet1"] = new SheetBinding
    {
        SheetName = "Sheet1",
        SystemKey = "current-business-system",
        ProjectId = "performance",
        ProjectName = "绩效项目",
        HeaderStartRow = 1,
        HeaderRowCount = 2,
        DataStartRow = 3,
    };

    var controller = CreateController(executionService, metadataStore, activeSheetName: "Sheet1");
    controller.RefreshActiveProjectFromSheetMetadata();

    controller.ExecuteInitializeCurrentSheet();

    Assert.Equal("Sheet1", executionService.LastInitializeSheetName);
}
```

- [ ] **Step 2: Run the Ribbon/controller tests**

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~RibbonSyncControllerTests`

Expected: FAIL because the controller still exposes incremental upload and no initialize action.

- [ ] **Step 3: Wire the new UI and controller behavior**

```csharp
public void SelectProject(ProjectOption project)
{
    var sheetName = GetRequiredSheetName();
    var binding = connector.CreateBindingSeed(sheetName, project);
    metadataStore.SaveBinding(binding);
    ApplyBindingState(binding);
    executionService.TryAutoInitializeCurrentSheet(sheetName, project);
}

public void ExecuteInitializeCurrentSheet()
{
    if (!EnsureProjectSelected())
    {
        return;
    }

    var sheetName = GetRequiredSheetName();
    executionService.InitializeCurrentSheet(sheetName, new ProjectOption
    {
        SystemKey = ActiveSystemKey,
        ProjectId = ActiveProjectId,
        DisplayName = ActiveProjectDisplayName,
    });
    OperationResultDialog.ShowInfo("初始化当前表完成。");
}
```

```csharp
private void InitializeSheetButton_Click(object sender, RibbonControlEventArgs e)
{
    Globals.ThisAddIn.RibbonSyncController?.ExecuteInitializeCurrentSheet();
}
```

```csharp
this.projectGroup.Items.Add(this.projectDropDown);
this.projectGroup.Items.Add(this.initializeSheetButton);
this.uploadGroup.Items.Remove(this.incrementalUploadButton);
```

```csharp
WorksheetSyncService = new WorksheetSyncService(
    CurrentBusinessConnector,
    WorksheetMetadataStore);
WorksheetSyncExecutionService = new WorksheetSyncExecutionService(
    WorksheetSyncService,
    WorksheetMetadataStore,
    new ExcelVisibleSelectionReader(Application),
    new ExcelWorksheetGridAdapter(Application),
    new SyncOperationPreviewFactory());
```

- [ ] **Step 4: Run the add-in tests and build the add-in project**

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~RibbonSyncControllerTests`

Expected: PASS

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj`

Expected: PASS, with no incremental-upload expectations left.

Run: `& "$env:ProgramFiles\\Microsoft Visual Studio\\2022\\Community\\MSBuild\\Current\\Bin\\MSBuild.exe" "src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj" /restore /p:RestorePackagesConfig=true /p:Configuration=Debug`

Expected: BUILD SUCCEEDED

- [ ] **Step 5: Commit**

```bash
git add src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs src/OfficeAgent.ExcelAddIn/AgentRibbon.cs src/OfficeAgent.ExcelAddIn/AgentRibbon.Designer.cs src/OfficeAgent.ExcelAddIn/ThisAddIn.cs tests/OfficeAgent.ExcelAddIn.Tests/RibbonSyncControllerTests.cs
git commit -m "feat: add ribbon sync initialize action"
```

### Task 7: Add Mock-Backed Coverage and Refresh Module Documentation

**Files:**
- Modify: `tests/mock-server/server.js`
- Modify: `tests/OfficeAgent.IntegrationTests/CurrentBusinessSystemConnectorIntegrationTests.cs`
- Modify: `docs/modules/ribbon-sync-current-behavior.md`
- Modify: `docs/module-index.md`
- Modify: `docs/vsto-manual-test-checklist.md`

- [ ] **Step 1: Write the failing integration test**

```csharp
[Fact]
public async Task BuildFieldMappingSeedAndBatchSaveRoundTripAgainstMockServer()
{
    var connector = await fixture.CreateCurrentBusinessConnector();

    var mappings = connector.BuildFieldMappingSeed("Sheet1", "performance");
    Assert.Contains(mappings, row => row.Values["ApiFieldKey"] == "start_12345678");

    var rows = connector.Find("performance", Array.Empty<string>(), new[] { "owner_name" });
    var rowId = rows.First()["row_id"].ToString();

    connector.BatchSave(
        "performance",
        new[]
        {
            new CellChange { RowId = rowId, ApiFieldKey = "owner_name", NewValue = "李四" },
        });

    var afterSave = connector.Find("performance", new[] { rowId }, new[] { "owner_name" });
    Assert.Equal("李四", afterSave.First()["owner_name"].ToString());
}
```

- [ ] **Step 2: Run the integration test to confirm the mock data shape still needs adjustment**

Run: `dotnet test tests/OfficeAgent.IntegrationTests/OfficeAgent.IntegrationTests.csproj --filter FullyQualifiedName~CurrentBusinessSystemConnectorIntegrationTests.BuildFieldMappingSeedAndBatchSaveRoundTripAgainstMockServer`

Expected: FAIL until the mock server returns the current-system head and row keys expected by the seed builder.

- [ ] **Step 3: Update the mock server and docs**

```javascript
apiApp.post('/head', requireAuth, function (_req, res) {
  return res.json({
    headList: [
      { fieldKey: 'row_id', headerText: 'ID', isId: true },
      { fieldKey: 'owner_name', headerText: '负责人' },
      { fieldKey: 'progress_status', headerText: '进展状态' },
      { headType: 'activity', activityId: '12345678', activityName: '测试活动111' },
    ],
  });
});

apiApp.post('/find', requireAuth, function (req, res) {
  var ids = ((req.body || {}).ids) || [];
  var fieldKeys = ((req.body || {}).fieldKeys) || [];
  var rows = performanceRows.filter(function (row) {
    return ids.length === 0 || ids.indexOf(row.row_id) >= 0;
  }).map(function (row) {
    if (fieldKeys.length === 0) return row;
    var projected = { row_id: row.row_id };
    fieldKeys.forEach(function (key) { projected[key] = row[key]; });
    return projected;
  });
  res.json(rows);
});
```

```markdown
| Ribbon Sync | [docs/modules/ribbon-sync-current-behavior.md](./modules/ribbon-sync-current-behavior.md) | [docs/superpowers/specs/2026-04-14-office-agent-ribbon-sync-configurability-design.md](./superpowers/specs/2026-04-14-office-agent-ribbon-sync-configurability-design.md)<br>[docs/superpowers/plans/2026-04-14-office-agent-ribbon-sync-configurability-implementation-plan.md](./superpowers/plans/2026-04-14-office-agent-ribbon-sync-configurability-implementation-plan.md) | [docs/vsto-manual-test-checklist.md](./vsto-manual-test-checklist.md)<br>[docs/ribbon-sync-real-system-integration-guide.md](./ribbon-sync-real-system-integration-guide.md)<br>[tests/mock-server/README.md](../tests/mock-server/README.md) |
```

```markdown
- 点击项目下拉框并选择项目，确认插件会自动尝试初始化当前 sheet。
- 在已有 Excel 上点击 `初始化当前表`，确认 `_OfficeAgentMetadata` 生成 `SheetBindings` 与 `SheetFieldMappings`。
- 将 `HeaderStartRow` 改为 `3`、`HeaderRowCount` 改为 `2`、`DataStartRow` 改为 `6` 后执行 `全量下载`，确认表头和数据写入新的行号位置。
- 修改 `SheetFieldMappings` 中的当前显示名后执行 `部分上传`，确认插件按当前表头文本匹配列而不是按旧列号。
- 确认 Ribbon 中不再出现 `增量上传` 按钮。
```

- [ ] **Step 4: Run the integration suite and full regression set**

Run: `dotnet test tests/OfficeAgent.IntegrationTests/OfficeAgent.IntegrationTests.csproj`

Expected: PASS

Run: `dotnet test tests/OfficeAgent.Core.Tests/OfficeAgent.Core.Tests.csproj`

Expected: PASS

Run: `dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj`

Expected: PASS

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj`

Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add tests/mock-server/server.js tests/OfficeAgent.IntegrationTests/CurrentBusinessSystemConnectorIntegrationTests.cs docs/modules/ribbon-sync-current-behavior.md docs/module-index.md docs/vsto-manual-test-checklist.md
git commit -m "test: cover configurable ribbon sync flow"
```

## Self-Review

- Spec coverage:
  - `SheetBindings` carries project binding plus row settings: Tasks 1 and 2
  - Dynamic `SheetFieldMappings` with semantic-role lookup: Tasks 1, 2, and 3
  - `HeaderStartRow`, `HeaderRowCount`, and `DataStartRow`: Tasks 1, 4, and 5
  - Existing Excel auto-try + explicit initialize fallback: Tasks 5 and 6
  - Header-text recognition instead of persisted column indexes: Tasks 4 and 5
  - Removal of incremental upload: Tasks 5 and 6
  - Connector and mock support for the current system: Tasks 3 and 7
  - Module snapshot and implementation entry docs: Task 7
- Placeholder scan:
  - No `TODO`, `TBD`, or “same as above” instructions remain
  - Each code-writing step includes concrete signatures or method bodies
- Type consistency:
  - `SheetBinding`, `FieldMappingTableDefinition`, `SheetFieldMappingRow`, `WorksheetRuntimeColumn`, `InitializeCurrentSheet`, and `BuildFieldMappingSeed` are used consistently across tasks
