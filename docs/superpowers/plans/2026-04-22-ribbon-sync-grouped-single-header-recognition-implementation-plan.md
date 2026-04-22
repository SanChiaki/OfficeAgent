# Ribbon Sync Grouped Single Header Recognition Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Support metadata-driven recognition of `single` fields displayed as two-row grouped Excel headers, without changing the default empty-sheet header generation style.

**Architecture:** Extend `WorksheetHeaderMatcher` so it builds three explicit lookup buckets: normal single, grouped single, and `activityProperty`, with early validation for duplicate two-level keys and unsupported `HeaderRowCount = 1` grouped-single matching. Keep default generation flat by changing `WorksheetSyncExecutionService.BuildConfiguredColumns(...)` to flatten grouped-single metadata to the child text when the sheet has no recognizable headers, while still reusing an already-recognized grouped layout for partial sync and existing-layout full download.

**Tech Stack:** C#, .NET Framework 4.8, VSTO Excel add-in, xUnit

---

**Assumption:** This plan assumes the previously approved `SheetFieldMappings` four-column display model (`ISDP L1`, `Excel L1`, `ISDP L2`, `Excel L2`) is already present on the branch.

## File Structure

- `src/OfficeAgent.ExcelAddIn/Excel/WorksheetHeaderMatcher.cs`
  Responsibility: classify mapping rows into normal single / grouped single / activity lookup buckets, reject conflicting two-level metadata, and match grouped single headers in two-row sheets.
- `src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs`
  Responsibility: keep grouped single support limited to runtime recognition for existing layouts, while flattening grouped single metadata to ordinary single display text on empty-sheet full download.
- `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetHeaderMatcherTests.cs`
  Responsibility: lock matcher behavior for grouped single success, normal-single regression coverage, duplicate-key validation, grouped-single/activity conflicts, and `HeaderRowCount = 1` rejection.
- `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs`
  Responsibility: lock partial download, partial upload, existing-layout full download, and empty-sheet full download behavior for grouped single metadata.
- `docs/modules/ribbon-sync-current-behavior.md`
  Responsibility: describe grouped single as a metadata-driven recognition capability, not a new `HeaderType` and not a new default generation mode.
- `docs/ribbon-sync-real-system-integration-guide.md`
  Responsibility: document the real-system contract for `single + Excel L2`, including the requirement that users maintain `Excel L1 / Excel L2` together with the visible Excel header.
- `docs/vsto-manual-test-checklist.md`
  Responsibility: add manual verification steps for grouped single partial sync, grouped single existing-layout full download, and empty-sheet flattening behavior.

### Task 1: Lock Grouped-Single Matching Rules

**Files:**
- Modify: `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetHeaderMatcherTests.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/Excel/WorksheetHeaderMatcher.cs`

- [ ] **Step 1: Add failing matcher tests for grouped single success and metadata errors**

```csharp
[Fact]
public void MatchUsesGroupedSingleHeadersForTwoRowHeaders()
{
    var grid = new FakeGrid();
    grid.SetCell("Sheet1", 3, 1, "ID");
    grid.SetCell("Sheet1", 3, 2, "联系人信息");
    grid.SetCell("Sheet1", 3, 3, "测试活动111");
    grid.SetCell("Sheet1", 4, 2, "负责人");
    grid.SetCell("Sheet1", 4, 3, "开始时间");

    var matcher = CreateMatcher();
    var binding = new SheetBinding
    {
        SheetName = "Sheet1",
        HeaderStartRow = 3,
        HeaderRowCount = 2,
    };
    var definition = BuildDefinition();
    var mappings = new[]
    {
        CreateMappingRow("Sheet1", "row_id", "single", true, currentSingle: "ID"),
        CreateMappingRow("Sheet1", "owner_name", "single", false, currentParent: "联系人信息", currentChild: "负责人"),
        CreateMappingRow("Sheet1", "start_12345678", "activityProperty", false, currentParent: "测试活动111", currentChild: "开始时间"),
    };

    var columns = InvokeMatch(matcher, "Sheet1", binding, definition, mappings, grid);

    Assert.Contains(columns, column =>
        column.ColumnIndex == 2 &&
        column.ApiFieldKey == "owner_name" &&
        string.Equals(column.HeaderType, "single", StringComparison.OrdinalIgnoreCase) &&
        string.Equals(column.DisplayText, "负责人", StringComparison.Ordinal));
}

[Fact]
public void MatchKeepsMergedSingleHeadersAheadOfTwoLevelMatching()
{
    var grid = new FakeGrid();
    grid.SetCell("Sheet1", 3, 1, "ID");
    grid.SetCell("Sheet1", 3, 2, "项目负责人");

    var matcher = CreateMatcher();
    var binding = new SheetBinding
    {
        SheetName = "Sheet1",
        HeaderStartRow = 3,
        HeaderRowCount = 2,
    };
    var definition = BuildDefinition();
    var mappings = new[]
    {
        CreateMappingRow("Sheet1", "row_id", "single", true, currentSingle: "ID"),
        CreateMappingRow("Sheet1", "owner_name", "single", false, currentSingle: "项目负责人"),
        CreateMappingRow("Sheet1", "owner_grouped", "single", false, currentParent: "联系人信息", currentChild: "负责人"),
    };

    var columns = InvokeMatch(matcher, "Sheet1", binding, definition, mappings, grid);

    Assert.Contains(columns, column => column.ColumnIndex == 2 && column.ApiFieldKey == "owner_name");
    Assert.DoesNotContain(columns, column => column.ApiFieldKey == "owner_grouped");
}

[Fact]
public void MatchThrowsWhenGroupedSingleKeyConflictsWithActivityKey()
{
    var matcher = CreateMatcher();
    var binding = new SheetBinding
    {
        SheetName = "Sheet1",
        HeaderStartRow = 3,
        HeaderRowCount = 2,
    };
    var definition = BuildDefinition();
    var mappings = new[]
    {
        CreateMappingRow("Sheet1", "owner_name", "single", false, currentParent: "联系人信息", currentChild: "负责人"),
        CreateMappingRow("Sheet1", "activity_owner", "activityProperty", false, currentParent: "联系人信息", currentChild: "负责人"),
    };

    var exception = Assert.Throws<TargetInvocationException>(() =>
        InvokeMatch(matcher, "Sheet1", binding, definition, mappings, new FakeGrid()));

    Assert.Equal(
        "SheetFieldMappings 中存在重复的双层表头键，请先修正 AI_Setting。",
        exception.InnerException?.Message);
}

[Fact]
public void MatchThrowsWhenGroupedSingleMetadataIsUsedWithSingleHeaderRow()
{
    var matcher = CreateMatcher();
    var binding = new SheetBinding
    {
        SheetName = "Sheet1",
        HeaderStartRow = 5,
        HeaderRowCount = 1,
    };
    var definition = BuildDefinition();
    var mappings = new[]
    {
        CreateMappingRow("Sheet1", "owner_name", "single", false, currentParent: "联系人信息", currentChild: "负责人"),
    };

    var exception = Assert.Throws<TargetInvocationException>(() =>
        InvokeMatch(matcher, "Sheet1", binding, definition, mappings, new FakeGrid()));

    Assert.Equal(
        "当前 HeaderRowCount=1，无法识别带 Excel L2 的 single 表头，请先修正 AI_Setting。",
        exception.InnerException?.Message);
}
```

- [ ] **Step 2: Run the matcher tests to verify they fail for the right reason**

Run:
```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~WorksheetHeaderMatcherTests
```

Expected:
- `MatchUsesGroupedSingleHeadersForTwoRowHeaders` fails because grouped single is not indexed or matched yet
- conflict/error tests fail because the matcher currently accepts duplicate two-level keys and grouped single under `HeaderRowCount = 1`

- [ ] **Step 3: Implement grouped-single lookup buckets and validation in `WorksheetHeaderMatcher`**

```csharp
public WorksheetRuntimeColumn[] Match(
    string sheetName,
    SheetBinding binding,
    FieldMappingTableDefinition definition,
    IReadOnlyList<SheetFieldMappingRow> mappings,
    IWorksheetGridAdapter grid)
{
    // ... existing guards ...
    var rows = mappings ?? Array.Empty<SheetFieldMappingRow>();
    var lookup = BuildLookup(definition, rows, binding.HeaderRowCount);
    // ... existing scan loop ...
}

private HeaderLookup BuildLookup(
    FieldMappingTableDefinition definition,
    IReadOnlyList<SheetFieldMappingRow> mappings,
    int headerRowCount)
{
    var singleHeaders = new Dictionary<string, WorksheetRuntimeColumn>(StringComparer.Ordinal);
    var groupedSingleHeaders = new Dictionary<string, WorksheetRuntimeColumn>(StringComparer.Ordinal);
    var activityHeaders = new Dictionary<string, WorksheetRuntimeColumn>(StringComparer.Ordinal);
    var twoLevelKeys = new HashSet<string>(StringComparer.Ordinal);
    var hasGroupedSingle = false;

    foreach (var mapping in mappings)
    {
        if (mapping == null)
        {
            continue;
        }

        var headerType = valueAccessor.GetValue(definition, mapping, FieldMappingSemanticRole.HeaderType);
        var apiFieldKey = valueAccessor.GetValue(definition, mapping, FieldMappingSemanticRole.ApiFieldKey);
        var currentSingle = valueAccessor.GetValue(definition, mapping, FieldMappingSemanticRole.CurrentSingleHeaderText);
        var currentParentText = valueAccessor.GetValue(definition, mapping, FieldMappingSemanticRole.CurrentParentHeaderText);
        var currentChildText = valueAccessor.GetValue(definition, mapping, FieldMappingSemanticRole.CurrentChildHeaderText);
        var isIdColumn = valueAccessor.GetBoolean(definition, mapping, FieldMappingSemanticRole.IsIdColumn);

        if (IsSingleHeader(headerType))
        {
            if (string.IsNullOrWhiteSpace(currentChildText))
            {
                if (!string.IsNullOrWhiteSpace(currentSingle) && !singleHeaders.ContainsKey(currentSingle))
                {
                    singleHeaders[currentSingle] = new WorksheetRuntimeColumn
                    {
                        ApiFieldKey = apiFieldKey,
                        HeaderType = string.IsNullOrWhiteSpace(headerType) ? "single" : headerType,
                        DisplayText = currentSingle,
                        ParentDisplayText = string.Empty,
                        ChildDisplayText = string.Empty,
                        IsIdColumn = isIdColumn,
                    };
                }
                continue;
            }

            hasGroupedSingle = true;
            var groupedKey = BuildTwoLevelKey(currentParentText, currentChildText);
            RegisterTwoLevelKey(twoLevelKeys, groupedKey);
            groupedSingleHeaders[groupedKey] = new WorksheetRuntimeColumn
            {
                ApiFieldKey = apiFieldKey,
                HeaderType = string.IsNullOrWhiteSpace(headerType) ? "single" : headerType,
                DisplayText = currentChildText,
                ParentDisplayText = currentParentText,
                ChildDisplayText = currentChildText,
                IsIdColumn = isIdColumn,
            };
            continue;
        }

        if (IsActivityProperty(headerType))
        {
            var activityKey = BuildTwoLevelKey(currentParentText, currentChildText);
            RegisterTwoLevelKey(twoLevelKeys, activityKey);
            activityHeaders[activityKey] = new WorksheetRuntimeColumn
            {
                ApiFieldKey = apiFieldKey,
                HeaderType = headerType,
                DisplayText = currentChildText,
                ParentDisplayText = currentParentText,
                ChildDisplayText = currentChildText,
                IsIdColumn = isIdColumn,
            };
        }
    }

    if (headerRowCount <= 1 && hasGroupedSingle)
    {
        throw new InvalidOperationException("当前 HeaderRowCount=1，无法识别带 Excel L2 的 single 表头，请先修正 AI_Setting。");
    }

    return new HeaderLookup(singleHeaders, groupedSingleHeaders, activityHeaders);
}

private WorksheetRuntimeColumn FindMatch(
    HeaderLookup lookup,
    string topText,
    string bottomText,
    string currentParent,
    int headerRowCount)
{
    if (headerRowCount <= 1)
    {
        return lookup.SingleHeaders.TryGetValue(topText, out var singleHeader)
            ? CloneSingleHeader(singleHeader)
            : null;
    }

    if (lookup.SingleHeaders.TryGetValue(topText, out var mergedSingleHeader) &&
        (string.IsNullOrWhiteSpace(bottomText) || string.Equals(bottomText, topText, StringComparison.Ordinal)))
    {
        return CloneSingleHeader(mergedSingleHeader);
    }

    var twoLevelKey = BuildTwoLevelKey(currentParent, bottomText);
    if (lookup.GroupedSingleHeaders.TryGetValue(twoLevelKey, out var groupedSingleHeader))
    {
        return new WorksheetRuntimeColumn
        {
            ApiFieldKey = groupedSingleHeader.ApiFieldKey,
            HeaderType = groupedSingleHeader.HeaderType,
            DisplayText = groupedSingleHeader.ChildDisplayText,
            ParentDisplayText = groupedSingleHeader.ParentDisplayText,
            ChildDisplayText = groupedSingleHeader.ChildDisplayText,
            IsIdColumn = groupedSingleHeader.IsIdColumn,
        };
    }

    return lookup.ActivityHeaders.TryGetValue(twoLevelKey, out var activityHeader)
        ? CloneActivityHeader(activityHeader)
        : null;
}

private static void RegisterTwoLevelKey(ISet<string> keys, string key)
{
    if (string.IsNullOrWhiteSpace(key))
    {
        return;
    }

    if (!keys.Add(key))
    {
        throw new InvalidOperationException("SheetFieldMappings 中存在重复的双层表头键，请先修正 AI_Setting。");
    }
}

private static string BuildTwoLevelKey(string parentText, string childText)
{
    if (string.IsNullOrWhiteSpace(parentText) || string.IsNullOrWhiteSpace(childText))
    {
        return string.Empty;
    }

    return parentText + "\u001f" + childText;
}

private sealed class HeaderLookup
{
    public HeaderLookup(
        IReadOnlyDictionary<string, WorksheetRuntimeColumn> singleHeaders,
        IReadOnlyDictionary<string, WorksheetRuntimeColumn> groupedSingleHeaders,
        IReadOnlyDictionary<string, WorksheetRuntimeColumn> activityHeaders)
    {
        SingleHeaders = singleHeaders ?? throw new ArgumentNullException(nameof(singleHeaders));
        GroupedSingleHeaders = groupedSingleHeaders ?? throw new ArgumentNullException(nameof(groupedSingleHeaders));
        ActivityHeaders = activityHeaders ?? throw new ArgumentNullException(nameof(activityHeaders));
    }

    public IReadOnlyDictionary<string, WorksheetRuntimeColumn> SingleHeaders { get; }

    public IReadOnlyDictionary<string, WorksheetRuntimeColumn> GroupedSingleHeaders { get; }

    public IReadOnlyDictionary<string, WorksheetRuntimeColumn> ActivityHeaders { get; }
}

private static bool IsActivityProperty(string headerType)
{
    return string.Equals(headerType, "activityProperty", StringComparison.OrdinalIgnoreCase);
}
```

- [ ] **Step 4: Run the matcher tests again to verify they pass**

Run:
```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~WorksheetHeaderMatcherTests
```

Expected:
- PASS

- [ ] **Step 5: Commit matcher support**

```bash
git add tests/OfficeAgent.ExcelAddIn.Tests/WorksheetHeaderMatcherTests.cs src/OfficeAgent.ExcelAddIn/Excel/WorksheetHeaderMatcher.cs
git commit -m "feat: recognize grouped single headers from metadata"
```

### Task 2: Support Grouped Single in Execution Paths Without Changing Default Layout Style

**Files:**
- Modify: `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs`

- [ ] **Step 1: Add failing execution-service tests for grouped single partial/full flows**

```csharp
[Fact]
public void ExecutePartialDownloadUsesGroupedSingleHeadersAndIdLookupOutsideSelection()
{
    var connector = new FakeSystemConnector();
    var metadataStore = new FakeWorksheetMetadataStore();
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
    metadataStore.Bindings["Sheet1"] = binding;
    metadataStore.FieldMappings["Sheet1"] = BuildGroupedSingleMappings("Sheet1");
    connector.FindResult = new[]
    {
        CreateRow("row-1", "张三", "2026-02-01", "2026-02-09"),
    };

    var selectionReader = new FakeWorksheetSelectionReader
    {
        VisibleCells = new[]
        {
            new SelectedVisibleCell { Row = 6, Column = 2, Value = "旧负责人" },
        },
    };
    var (service, grid) = CreateService(connector, metadataStore, selectionReader);
    SeedGroupedSingleHeaders(grid, "Sheet1", binding);
    grid.SetCell("Sheet1", 6, 1, "row-1");
    grid.SetCell("Sheet1", 6, 2, "旧负责人");

    var plan = InvokePrepare(service, "PreparePartialDownload", "Sheet1");
    InvokeExecute(service, "ExecuteDownload", plan);

    Assert.Equal(new[] { "owner_name" }, connector.LastFindFieldKeys);
    Assert.Equal("张三", grid.GetCell("Sheet1", 6, 2));
}

[Fact]
public void ExecutePartialUploadUsesGroupedSingleHeadersAndIdLookupOutsideSelection()
{
    var connector = new FakeSystemConnector();
    var metadataStore = new FakeWorksheetMetadataStore();
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
    metadataStore.Bindings["Sheet1"] = binding;
    metadataStore.FieldMappings["Sheet1"] = BuildGroupedSingleMappings("Sheet1");

    var selectionReader = new FakeWorksheetSelectionReader
    {
        VisibleCells = new[]
        {
            new SelectedVisibleCell { Row = 6, Column = 2, Value = "李四" },
        },
    };
    var (service, grid) = CreateService(connector, metadataStore, selectionReader);
    SeedGroupedSingleHeaders(grid, "Sheet1", binding);
    grid.SetCell("Sheet1", 6, 1, "row-1");
    grid.SetCell("Sheet1", 6, 2, "李四");

    var plan = InvokePrepare(service, "PreparePartialUpload", "Sheet1");
    var preview = ReadPreview(plan);

    Assert.Single(preview.Changes);
    Assert.Equal("row-1", preview.Changes[0].RowId);
    Assert.Equal("owner_name", preview.Changes[0].ApiFieldKey);
}

[Fact]
public void PrepareFullDownloadReusesExistingGroupedSingleLayout()
{
    var connector = new FakeSystemConnector();
    var metadataStore = new FakeWorksheetMetadataStore();
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
    metadataStore.Bindings["Sheet1"] = binding;
    metadataStore.FieldMappings["Sheet1"] = BuildGroupedSingleMappings("Sheet1");
    connector.FindResult = new[]
    {
        CreateRow("row-1", "张三", "2026-02-01", "2026-02-09"),
    };

    var (service, grid) = CreateService(connector, metadataStore, new FakeWorksheetSelectionReader());
    SeedGroupedSingleHeaders(grid, "Sheet1", binding);

    var plan = InvokePrepare(service, "PrepareFullDownload", "Sheet1");

    Assert.True(ReadUsesExistingLayout(plan));
}

[Fact]
public void ExecuteFullDownloadFlattensGroupedSingleHeadersWhenSheetHeadersAreEmpty()
{
    var connector = new FakeSystemConnector();
    var metadataStore = new FakeWorksheetMetadataStore();
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
    metadataStore.Bindings["Sheet1"] = binding;
    metadataStore.FieldMappings["Sheet1"] = BuildGroupedSingleMappings("Sheet1");
    connector.FindResult = new[]
    {
        CreateRow("row-1", "张三", "2026-02-01", "2026-02-09"),
    };

    var (service, grid) = CreateService(connector, metadataStore, new FakeWorksheetSelectionReader());

    var plan = InvokePrepare(service, "PrepareFullDownload", "Sheet1");
    InvokeExecute(service, "ExecuteDownload", plan);

    Assert.False(ReadUsesExistingLayout(plan));
    Assert.Equal("负责人", grid.GetCell("Sheet1", 3, 2));
    Assert.Equal(string.Empty, grid.GetCell("Sheet1", 4, 2));
    Assert.Contains(grid.Merges, merge => merge.SheetName == "Sheet1" && merge.Row == 3 && merge.Column == 2 && merge.RowSpan == 2 && merge.ColumnSpan == 1);
}

private static bool ReadUsesExistingLayout(object plan)
{
    var property = plan.GetType().GetProperty(
        "UsesExistingLayout",
        BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

    if (property == null)
    {
        throw new InvalidOperationException("UsesExistingLayout property was not found.");
    }

    return (bool)property.GetValue(plan);
}

private static SheetFieldMappingRow[] BuildGroupedSingleMappings(string sheetName)
{
    return new[]
    {
        CreateMappingRow(sheetName, "row_id", "single", true, currentSingle: "ID"),
        CreateMappingRow(
            sheetName,
            "owner_name",
            "single",
            false,
            defaultSingle: "负责人",
            currentParent: "联系人信息",
            currentChild: "负责人"),
        CreateMappingRow(
            sheetName,
            "start_12345678",
            "activityProperty",
            false,
            defaultParent: "测试活动111",
            currentParent: "测试活动111",
            defaultChild: "开始时间",
            currentChild: "开始时间",
            activityId: "12345678",
            propertyId: "start"),
        CreateMappingRow(
            sheetName,
            "end_12345678",
            "activityProperty",
            false,
            defaultParent: "测试活动111",
            currentParent: "测试活动111",
            defaultChild: "结束时间",
            currentChild: "结束时间",
            activityId: "12345678",
            propertyId: "end"),
    };
}

private static void SeedGroupedSingleHeaders(FakeWorksheetGridAdapter grid, string sheetName, SheetBinding binding)
{
    var row = binding.HeaderStartRow;
    grid.SetCell(sheetName, row, 1, "ID");
    grid.SetCell(sheetName, row, 2, "联系人信息");
    grid.SetCell(sheetName, row, 3, "测试活动111");

    if (binding.HeaderRowCount > 1)
    {
        grid.SetCell(sheetName, row + 1, 2, "负责人");
        grid.SetCell(sheetName, row + 1, 3, "开始时间");
        grid.SetCell(sheetName, row + 1, 4, "结束时间");
    }
}
```

- [ ] **Step 2: Run the grouped-single execution tests to verify they fail first**

Run:
```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~ExecutePartialDownloadUsesGroupedSingleHeadersAndIdLookupOutsideSelection|FullyQualifiedName~ExecutePartialUploadUsesGroupedSingleHeadersAndIdLookupOutsideSelection|FullyQualifiedName~PrepareFullDownloadReusesExistingGroupedSingleLayout|FullyQualifiedName~ExecuteFullDownloadFlattensGroupedSingleHeadersWhenSheetHeadersAreEmpty"
```

Expected:
- partial download/upload tests fail because grouped single columns are not resolved yet
- empty-sheet full download test fails because `BuildConfiguredColumns(...)` still uses `Excel L1` for `single`, not `Excel L2`

- [ ] **Step 3: Flatten grouped single only in the empty-sheet configured-column path**

```csharp
private WorksheetRuntimeColumn[] BuildConfiguredColumns(
    SheetBinding binding,
    FieldMappingTableDefinition definition,
    IReadOnlyList<SheetFieldMappingRow> mappings)
{
    var rows = mappings ?? Array.Empty<SheetFieldMappingRow>();
    var result = new List<WorksheetRuntimeColumn>(rows.Count);
    var columnIndex = 1;

    foreach (var mapping in rows)
    {
        if (mapping == null)
        {
            continue;
        }

        var apiFieldKey = valueAccessor.GetValue(definition, mapping, FieldMappingSemanticRole.ApiFieldKey);
        if (string.IsNullOrWhiteSpace(apiFieldKey))
        {
            continue;
        }

        var headerType = valueAccessor.GetValue(definition, mapping, FieldMappingSemanticRole.HeaderType);
        var isActivityProperty = IsActivityProperty(headerType);
        var singleText = valueAccessor.GetValue(definition, mapping, FieldMappingSemanticRole.CurrentSingleHeaderText);
        var parentText = valueAccessor.GetValue(definition, mapping, FieldMappingSemanticRole.CurrentParentHeaderText);
        var childText = valueAccessor.GetValue(definition, mapping, FieldMappingSemanticRole.CurrentChildHeaderText);
        var displayText = isActivityProperty
            ? childText
            : ResolveConfiguredSingleDisplayText(singleText, childText);

        result.Add(new WorksheetRuntimeColumn
        {
            ColumnIndex = columnIndex++,
            ApiFieldKey = apiFieldKey,
            HeaderType = NormalizeHeaderType(headerType),
            DisplayText = displayText,
            ParentDisplayText = isActivityProperty && binding.HeaderRowCount > 1 ? parentText : string.Empty,
            ChildDisplayText = isActivityProperty ? childText : string.Empty,
            IsIdColumn = valueAccessor.GetBoolean(definition, mapping, FieldMappingSemanticRole.IsIdColumn),
        });
    }

    return result.ToArray();
}

private static string ResolveConfiguredSingleDisplayText(string singleText, string childText)
{
    return string.IsNullOrWhiteSpace(childText)
        ? singleText
        : childText;
}
```

Implementation notes:
- do not change `WorksheetSchemaLayoutService`; it should keep rendering ordinary singles from `WorksheetRuntimeColumn.DisplayText`
- do not invent a new `HeaderType`; grouped single continues to flow as `single`

- [ ] **Step 4: Run the grouped-single execution tests, then the full add-in suite**

Run:
```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~ExecutePartialDownloadUsesGroupedSingleHeadersAndIdLookupOutsideSelection|FullyQualifiedName~ExecutePartialUploadUsesGroupedSingleHeadersAndIdLookupOutsideSelection|FullyQualifiedName~PrepareFullDownloadReusesExistingGroupedSingleLayout|FullyQualifiedName~ExecuteFullDownloadFlattensGroupedSingleHeadersWhenSheetHeadersAreEmpty"
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj
```

Expected:
- focused grouped-single tests PASS
- full `OfficeAgent.ExcelAddIn.Tests` suite PASS

- [ ] **Step 5: Commit execution-path support**

```bash
git add tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs
git commit -m "feat: support grouped single sync layouts"
```

### Task 3: Update Docs and Manual Verification Guidance

**Files:**
- Modify: `docs/modules/ribbon-sync-current-behavior.md`
- Modify: `docs/ribbon-sync-real-system-integration-guide.md`
- Modify: `docs/vsto-manual-test-checklist.md`

- [ ] **Step 1: Update the module snapshot and integration guide with grouped-single rules**

Add the following points to `docs/modules/ribbon-sync-current-behavior.md`:

```markdown
- `HeaderType = single` 且 `Excel L2` 为空时，按普通 single 匹配
- `HeaderType = single` 且 `Excel L2` 非空时，按“分组 single”处理，但字段身份仍然是 `single`
- `部分下载`、`部分上传`、以及“已有可识别表头”的 `全量下载` 支持这种 metadata 驱动识别
- 空表头时，插件仍按当前默认 single 样式生成表头，不会主动生成分组父头
- `HeaderRowCount = 1` 时，带 `Excel L2` 的 `single` 会直接报 `AI_Setting` 配置错误
```

Add the following points to `docs/ribbon-sync-real-system-integration-guide.md`:

```markdown
- `single + Excel L2` 表示“分组 single”，只能用于当前表头识别，不代表连接器默认要生成双层 single 布局
- 如果用户把单字段改成两层 Excel 表头，必须同步维护 `SheetFieldMappings.Excel L1 / Excel L2`
- 只改 Excel、不改 metadata 时，插件不会做自动猜测识别
- 空表头场景下，默认生成仍然取字段名本身，不会自动补一层分组父头
```

- [ ] **Step 2: Extend the manual checklist with grouped-single smoke tests**

Append checklist items like:

```markdown
- 把某个 `single` 字段的 `Excel L1 / Excel L2` 改成 `联系人信息 / 负责人`，并把业务表头手工改成两行分组表头，然后执行 `部分下载`，确认该列仍按 `owner_name` 命中。
- 在同样的 metadata 和 Excel 布局下执行 `部分上传`，确认提交的 `ApiFieldKey` 仍是该 `single` 字段，而不是活动字段。
- 保留上述已识别的双层 grouped single 表头后执行 `全量下载`，确认插件复用现有布局，不重写表头。
- 清空表头区后再次执行 `全量下载`，确认插件重新生成的是平铺 single 文本 `负责人`，不会自动写出 `联系人信息` 这一层父头。
```

- [ ] **Step 3: Run focused regression tests after the doc-linked implementation lands**

Run:
```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~WorksheetHeaderMatcherTests|FullyQualifiedName~WorksheetSyncExecutionServiceTests"
```

Expected:
- PASS

- [ ] **Step 4: Refresh the local add-in and do one manual Excel smoke pass**

Run:
```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File eng/Dev-RefreshExcelAddIn.ps1 -CloseExcel
```

Expected:
- frontend bundle and Debug add-in rebuild successfully
- Excel registration refresh succeeds

Then manually verify one sheet with:
- `AI_Setting` grouped single metadata (`Excel L1 = 联系人信息`, `Excel L2 = 负责人`)
- existing two-row grouped headers for partial sync / existing-layout full download
- empty header area for flattening verification

- [ ] **Step 5: Commit docs and checklist updates**

```bash
git add docs/modules/ribbon-sync-current-behavior.md docs/ribbon-sync-real-system-integration-guide.md docs/vsto-manual-test-checklist.md
git commit -m "docs: describe grouped single header recognition"
```

## Self-Review

- Spec coverage:
  - `single + Excel L2` stays `HeaderType = single`: Task 1 and Task 2
  - matcher keeps separate normal single / grouped single / activity buckets: Task 1
  - conflict detection for duplicate two-level keys: Task 1
  - `HeaderRowCount = 1` grouped single rejection: Task 1
  - support only partial download / partial upload / existing-layout full download: Task 2
  - empty-sheet full download stays flat and does not generate grouped parent headers: Task 2
  - docs and manual verification updated: Task 3
- Placeholder scan:
  - no placeholder markers remain
  - each code-changing step includes concrete test names, implementation snippets, commands, and expected outcomes
- Type consistency:
  - `WorksheetHeaderMatcher`, `BuildConfiguredColumns`, `ResolveConfiguredSingleDisplayText`, `BuildGroupedSingleMappings`, and `ReadUsesExistingLayout` are named consistently across tasks
