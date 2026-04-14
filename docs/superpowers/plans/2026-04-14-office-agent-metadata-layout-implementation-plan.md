# OfficeAgent Metadata Layout Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Rebuild `_OfficeAgentMetadata` into a human-readable worksheet layout with two titled table sections while preserving the existing Ribbon Sync metadata model.

**Architecture:** Keep `SheetBindings` and `SheetFieldMappings` as the only logical metadata tables, but stop storing them as compressed `[tableName, values...]` rows. Instead, parse and render the sheet as two explicit sections: a title row, a header row, then data rows. This is a display/storage change only; upload/download semantics and metadata meaning stay unchanged.

**Tech Stack:** C#, .NET Framework 4.8, VSTO, Excel Interop, xUnit

---

## File Structure

- `src/OfficeAgent.ExcelAddIn/Excel/MetadataSectionDocument.cs`
  Responsibility: represent one titled metadata section and convert between section rows and logical table rows.
- `src/OfficeAgent.ExcelAddIn/Excel/MetadataSheetLayoutSerializer.cs`
  Responsibility: render and parse the full `_OfficeAgentMetadata` sheet as two canonical sections in fixed order.
- `src/OfficeAgent.ExcelAddIn/Excel/ExcelWorkbookMetadataAdapter.cs`
  Responsibility: read/write `_OfficeAgentMetadata` by loading the full sheet, replacing one logical section, and rewriting the sheet in the new readable layout while preserving worksheet focus.
- `tests/OfficeAgent.ExcelAddIn.Tests/MetadataSheetLayoutSerializerTests.cs`
  Responsibility: verify section rendering, parsing, blank-row separators, and missing-table behavior without COM fakes.
- `tests/OfficeAgent.ExcelAddIn.Tests/ExcelWorkbookMetadataAdapterTests.cs`
  Responsibility: verify adapter still restores focus and rewrites metadata as readable titled sections.
- `docs/modules/ribbon-sync-current-behavior.md`
  Responsibility: describe the new `_OfficeAgentMetadata` readable layout.
- `docs/vsto-manual-test-checklist.md`
  Responsibility: add manual verification steps for the new metadata display.

### Task 1: Add a Serializer for Readable Metadata Sections

**Files:**
- Create: `src/OfficeAgent.ExcelAddIn/Excel/MetadataSectionDocument.cs`
- Create: `src/OfficeAgent.ExcelAddIn/Excel/MetadataSheetLayoutSerializer.cs`
- Create: `tests/OfficeAgent.ExcelAddIn.Tests/MetadataSheetLayoutSerializerTests.cs`

- [ ] **Step 1: Write the failing serializer tests**

```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class MetadataSheetLayoutSerializerTests
    {
        [Fact]
        public void RenderPlacesBindingsAboveFieldMappingsUsingReadableSections()
        {
            var serializer = CreateSerializer();
            var rendered = InvokeRender(
                serializer,
                new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["SheetBindings"] = CreateSection(
                        "SheetBindings",
                        new[] { "SheetName", "SystemKey" },
                        new[] { new[] { "Sheet1", "current-business-system" } }),
                    ["SheetFieldMappings"] = CreateSection(
                        "SheetFieldMappings",
                        new[] { "SheetName", "HeaderId", "ApiFieldKey" },
                        new[] { new[] { "Sheet1", "row_id", "row_id" } }),
                });

            Assert.Equal("SheetBindings", rendered[0][0]);
            Assert.Equal(new[] { "SheetName", "SystemKey" }, rendered[1]);
            Assert.Equal(new[] { "Sheet1", "current-business-system" }, rendered[2]);
            Assert.True(rendered[3].All(string.IsNullOrEmpty));
            Assert.True(rendered[4].All(string.IsNullOrEmpty));
            Assert.Equal("SheetFieldMappings", rendered[5][0]);
            Assert.Equal(new[] { "SheetName", "HeaderId", "ApiFieldKey" }, rendered[6]);
            Assert.Equal(new[] { "Sheet1", "row_id", "row_id" }, rendered[7]);
        }

        [Fact]
        public void ReadTableReturnsOnlySectionDataRows()
        {
            var serializer = CreateSerializer();
            var sheetRows = new[]
            {
                new[] { "SheetBindings" },
                new[] { "SheetName", "SystemKey", "ProjectId" },
                new[] { "Sheet1", "current-business-system", "performance" },
                new[] { string.Empty, string.Empty, string.Empty },
                new[] { string.Empty, string.Empty, string.Empty },
                new[] { "SheetFieldMappings" },
                new[] { "SheetName", "HeaderId", "ApiFieldKey", "IsIdColumn" },
                new[] { "Sheet1", "row_id", "row_id", "TRUE" },
                new[] { "Sheet1", "owner_name", "owner_name", "FALSE" },
            };

            var rows = InvokeReadTable(serializer, "SheetFieldMappings", sheetRows);

            Assert.Equal(2, rows.Length);
            Assert.Equal(new[] { "Sheet1", "row_id", "row_id", "TRUE" }, rows[0]);
            Assert.Equal(new[] { "Sheet1", "owner_name", "owner_name", "FALSE" }, rows[1]);
        }

        [Fact]
        public void ReadTableReturnsEmptyWhenSectionMissing()
        {
            var serializer = CreateSerializer();

            var rows = InvokeReadTable(serializer, "SheetFieldMappings", new[]
            {
                new[] { "SheetBindings" },
                new[] { "SheetName", "SystemKey" },
                new[] { "Sheet1", "current-business-system" },
            });

            Assert.Empty(rows);
        }

        private static object CreateSerializer()
        {
            var assembly = LoadAddInAssembly();
            var serializerType = assembly.GetType("OfficeAgent.ExcelAddIn.Excel.MetadataSheetLayoutSerializer", throwOnError: true);
            return Activator.CreateInstance(serializerType);
        }

        private static object CreateSection(string title, string[] headers, string[][] rows)
        {
            var assembly = LoadAddInAssembly();
            var type = assembly.GetType("OfficeAgent.ExcelAddIn.Excel.MetadataSectionDocument", throwOnError: true);
            return Activator.CreateInstance(type, title, headers, rows);
        }

        private static string[][] InvokeRender(object serializer, IDictionary<string, object> sections)
        {
            var method = serializer.GetType().GetMethod("Render", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            return (string[][])method.Invoke(serializer, new object[] { sections });
        }

        private static string[][] InvokeReadTable(object serializer, string tableName, string[][] sheetRows)
        {
            var method = serializer.GetType().GetMethod("ReadTable", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            return (string[][])method.Invoke(serializer, new object[] { tableName, sheetRows });
        }

        private static Assembly LoadAddInAssembly()
        {
            return Assembly.LoadFrom(System.IO.Path.GetFullPath(
                System.IO.Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "src", "OfficeAgent.ExcelAddIn", "bin", "Debug", "OfficeAgent.ExcelAddIn.dll")));
        }
    }
}
```

- [ ] **Step 2: Run the serializer tests to verify they fail**

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~MetadataSheetLayoutSerializerTests`

Expected: FAIL with type-not-found errors for `MetadataSectionDocument` and `MetadataSheetLayoutSerializer`.

- [ ] **Step 3: Add the minimal serializer types**

```csharp
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal sealed class MetadataSectionDocument
    {
        public MetadataSectionDocument(string title, string[] headers, string[][] rows)
        {
            Title = title ?? string.Empty;
            Headers = headers ?? Array.Empty<string>();
            Rows = rows ?? Array.Empty<string[]>();
        }

        public string Title { get; }

        public string[] Headers { get; }

        public string[][] Rows { get; }
    }

    internal sealed class MetadataSheetLayoutSerializer
    {
        private static readonly string[] SectionOrder = { "SheetBindings", "SheetFieldMappings" };

        public string[][] Render(IReadOnlyDictionary<string, MetadataSectionDocument> sections)
        {
            var rendered = new List<string[]>();

            foreach (var title in SectionOrder)
            {
                if (sections == null || !sections.TryGetValue(title, out var section) || section == null)
                {
                    continue;
                }

                if (rendered.Count > 0)
                {
                    rendered.Add(Array.Empty<string>());
                    rendered.Add(Array.Empty<string>());
                }

                rendered.Add(new[] { section.Title });
                rendered.Add(section.Headers ?? Array.Empty<string>());
                rendered.AddRange(section.Rows ?? Array.Empty<string[]>());
            }

            return rendered.ToArray();
        }

        public string[][] ReadTable(string tableName, string[][] sheetRows)
        {
            if (string.IsNullOrWhiteSpace(tableName) || sheetRows == null)
            {
                return Array.Empty<string[]>();
            }

            for (var rowIndex = 0; rowIndex < sheetRows.Length; rowIndex++)
            {
                var row = sheetRows[rowIndex] ?? Array.Empty<string>();
                if (row.Length == 0 || !string.Equals(row[0], tableName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var result = new List<string[]>();
                for (var dataRow = rowIndex + 2; dataRow < sheetRows.Length; dataRow++)
                {
                    var candidate = sheetRows[dataRow] ?? Array.Empty<string>();
                    if (candidate.Length > 0 && !string.IsNullOrWhiteSpace(candidate[0]) && Array.IndexOf(SectionOrder, candidate[0]) >= 0)
                    {
                        break;
                    }

                    if (candidate.All(string.IsNullOrEmpty))
                    {
                        break;
                    }

                    result.Add(candidate);
                }

                return result.ToArray();
            }

            return Array.Empty<string[]>();
        }
    }
}
```

- [ ] **Step 4: Run the serializer tests to verify they pass**

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~MetadataSheetLayoutSerializerTests`

Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/OfficeAgent.ExcelAddIn/Excel/MetadataSectionDocument.cs src/OfficeAgent.ExcelAddIn/Excel/MetadataSheetLayoutSerializer.cs tests/OfficeAgent.ExcelAddIn.Tests/MetadataSheetLayoutSerializerTests.cs
git commit -m "test: add readable metadata layout serializer"
```

### Task 2: Rewrite the Excel Metadata Adapter to Use Readable Sections

**Files:**
- Modify: `src/OfficeAgent.ExcelAddIn/Excel/ExcelWorkbookMetadataAdapter.cs`
- Modify: `tests/OfficeAgent.ExcelAddIn.Tests/ExcelWorkbookMetadataAdapterTests.cs`

- [ ] **Step 1: Write the failing adapter tests**

```csharp
[Fact]
public void WriteTableRewritesMetadataSheetUsingTitledSections()
{
    var addInAssembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
    var excelAssembly = LoadExcelInteropAssembly();
    var worksheetType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Worksheet", throwOnError: true);
    var sheetsType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Sheets", throwOnError: true);
    var workbookType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Workbook", throwOnError: true);
    var applicationType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Application", throwOnError: true);

    var application = new LayoutAwareFakeExcelApplication(applicationType, workbookType, sheetsType, worksheetType);
    var adapterType = addInAssembly.GetType("OfficeAgent.ExcelAddIn.Excel.ExcelWorkbookMetadataAdapter", throwOnError: true);
    var adapter = Activator.CreateInstance(adapterType, application.GetTransparentProxy());

    adapterType.GetMethod("WriteTable").Invoke(adapter, new object[]
    {
        "SheetBindings",
        new[] { "SheetName", "SystemKey" },
        new[] { new[] { "Sheet1", "current-business-system" } },
    });

    adapterType.GetMethod("WriteTable").Invoke(adapter, new object[]
    {
        "SheetFieldMappings",
        new[] { "SheetName", "HeaderId", "ApiFieldKey" },
        new[] { new[] { "Sheet1", "row_id", "row_id" } },
    });

    Assert.Equal("SheetBindings", application.MetadataSheet.GetCell(1, 1));
    Assert.Equal("SheetName", application.MetadataSheet.GetCell(2, 1));
    Assert.Equal("Sheet1", application.MetadataSheet.GetCell(3, 1));
    Assert.Equal("SheetFieldMappings", application.MetadataSheet.GetCell(6, 1));
    Assert.Equal("HeaderId", application.MetadataSheet.GetCell(7, 2));
    Assert.Equal("row_id", application.MetadataSheet.GetCell(8, 2));
}

[Fact]
public void ReadTableReadsRowsBackFromTitledSections()
{
    var addInAssembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
    var excelAssembly = LoadExcelInteropAssembly();
    var worksheetType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Worksheet", throwOnError: true);
    var sheetsType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Sheets", throwOnError: true);
    var workbookType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Workbook", throwOnError: true);
    var applicationType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Application", throwOnError: true);

    var application = new LayoutAwareFakeExcelApplication(applicationType, workbookType, sheetsType, worksheetType);
    application.MetadataSheet.SetCell(1, 1, "SheetBindings");
    application.MetadataSheet.SetCell(2, 1, "SheetName");
    application.MetadataSheet.SetCell(2, 2, "SystemKey");
    application.MetadataSheet.SetCell(3, 1, "Sheet1");
    application.MetadataSheet.SetCell(3, 2, "current-business-system");

    var adapterType = addInAssembly.GetType("OfficeAgent.ExcelAddIn.Excel.ExcelWorkbookMetadataAdapter", throwOnError: true);
    var adapter = Activator.CreateInstance(adapterType, application.GetTransparentProxy());

    var rows = (string[][])adapterType.GetMethod("ReadTable").Invoke(adapter, new object[] { "SheetBindings" });

    Assert.Single(rows);
    Assert.Equal(new[] { "Sheet1", "current-business-system" }, rows[0]);
}
```

- [ ] **Step 2: Run the adapter tests to verify they fail**

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~ExcelWorkbookMetadataAdapterTests`

Expected: FAIL because the current adapter still writes compressed rows using the table name in the first column.

- [ ] **Step 3: Rework the adapter to read and write canonical sections**

```csharp
private static readonly string[] OrderedTables = { "SheetBindings", "SheetFieldMappings" };
private readonly MetadataSheetLayoutSerializer serializer = new MetadataSheetLayoutSerializer();

public void WriteTable(string tableName, string[] headers, string[][] rows)
{
    ExecutePreservingActiveWorksheet(() =>
    {
        var worksheet = EnsureWorksheetExists(MetadataSheetName);
        var sections = LoadAllSections(worksheet);
        sections[tableName] = new MetadataSectionDocument(tableName, headers, rows);
        RewriteSheet(worksheet, sections);
    });
}

public string[][] ReadTable(string tableName)
{
    return ExecutePreservingActiveWorksheet(() =>
    {
        var worksheet = EnsureWorksheetExists(MetadataSheetName);
        var sheetRows = ReadUsedRows(worksheet);
        return serializer.ReadTable(tableName, sheetRows);
    });
}

private Dictionary<string, MetadataSectionDocument> LoadAllSections(ExcelInterop.Worksheet worksheet)
{
    var sheetRows = ReadUsedRows(worksheet);
    var sections = new Dictionary<string, MetadataSectionDocument>(StringComparer.OrdinalIgnoreCase);

    foreach (var tableName in OrderedTables)
    {
        var rows = serializer.ReadTable(tableName, sheetRows);
        var headers = serializer.ReadHeaders(tableName, sheetRows);
        if (headers.Length == 0)
        {
            continue;
        }

        sections[tableName] = new MetadataSectionDocument(tableName, headers, rows);
    }

    return sections;
}

private void RewriteSheet(ExcelInterop.Worksheet worksheet, IReadOnlyDictionary<string, MetadataSectionDocument> sections)
{
    worksheet.Cells.ClearContents();
    var rendered = serializer.Render(sections);
    for (var row = 0; row < rendered.Length; row++)
    {
        var values = rendered[row] ?? Array.Empty<string>();
        for (var column = 0; column < values.Length; column++)
        {
            (worksheet.Cells[row + 1, column + 1] as ExcelInterop.Range).Value2 = values[column];
        }
    }
}
```

- [ ] **Step 4: Run the adapter tests to verify they pass**

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~ExcelWorkbookMetadataAdapterTests`

Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/OfficeAgent.ExcelAddIn/Excel/ExcelWorkbookMetadataAdapter.cs tests/OfficeAgent.ExcelAddIn.Tests/ExcelWorkbookMetadataAdapterTests.cs
git commit -m "feat: render metadata sheet as readable sections"
```

### Task 3: Refresh Store-Level Verification and User-Facing Documentation

**Files:**
- Modify: `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetMetadataStoreTests.cs`
- Modify: `docs/modules/ribbon-sync-current-behavior.md`
- Modify: `docs/vsto-manual-test-checklist.md`

- [ ] **Step 1: Write the failing store/doc-oriented tests**

```csharp
[Fact]
public void SaveFieldMappingsUsesSheetNameAsAVisibleBusinessColumn()
{
    var (store, adapter) = CreateStore();
    var definition = new FieldMappingTableDefinition
    {
        SystemKey = "current-business-system",
        Columns = new[]
        {
            new FieldMappingColumnDefinition { ColumnName = "HeaderId", Role = FieldMappingSemanticRole.HeaderIdentity },
            new FieldMappingColumnDefinition { ColumnName = "ApiFieldKey", Role = FieldMappingSemanticRole.ApiFieldKey },
        },
    };

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
                    ["HeaderId"] = "row_id",
                    ["ApiFieldKey"] = "row_id",
                },
            },
        });

    var loaded = InvokeLoadFieldMappings(store, "Sheet1", definition);

    Assert.Single(loaded);
    Assert.Equal("row_id", loaded[0].Values["HeaderId"]);
    Assert.Equal("row_id", loaded[0].Values["ApiFieldKey"]);
}
```

- [ ] **Step 2: Run the focused store tests**

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~WorksheetMetadataStoreTests`

Expected: PASS once the adapter rewrite does not break logical table roundtrips.

- [ ] **Step 3: Update the docs to describe the new metadata layout**

```markdown
## 元数据模型

`_OfficeAgentMetadata` 当前采用同一个 sheet 内上下两个区域：

- `SheetBindings`
- `SheetFieldMappings`

每个区域都包含：

- 一行区域标题
- 一行表头
- 多行数据
```

```markdown
- 打开 `_OfficeAgentMetadata`，确认 `SheetBindings` 位于上方，`SheetFieldMappings` 位于下方。
- 确认每个区域都有标题行、表头行、数据区，而不是旧的压平行格式。
- 手工修改 `SheetBindings.HeaderStartRow` 后再次执行下载，确认插件继续按更新后的值工作。
```

- [ ] **Step 4: Run the add-in suite and spot-check docs**

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj`

Expected: PASS

Run: `git diff -- docs/modules/ribbon-sync-current-behavior.md docs/vsto-manual-test-checklist.md`

Expected: shows only metadata-layout documentation updates.

- [ ] **Step 5: Commit**

```bash
git add tests/OfficeAgent.ExcelAddIn.Tests/WorksheetMetadataStoreTests.cs docs/modules/ribbon-sync-current-behavior.md docs/vsto-manual-test-checklist.md
git commit -m "docs: describe readable metadata sheet layout"
```

## Self-Review

- Spec coverage:
  - `_OfficeAgentMetadata` stays as one sheet: Tasks 1 and 2
  - `SheetBindings` above `SheetFieldMappings`: Tasks 1 and 2
  - Each section uses title row + header row + data rows: Tasks 1 and 2
  - No old-format compatibility: Task 2 rewrite clears and rewrites the metadata sheet in canonical layout
  - Human-readable debugging and manual maintenance: Tasks 2 and 3
- Placeholder scan:
  - No `TODO`, `TBD`, or “same as above” references remain
  - Each code-changing step includes concrete test or implementation code
- Type consistency:
  - `MetadataSectionDocument`, `MetadataSheetLayoutSerializer`, `WriteTable`, and `ReadTable` are used consistently across tasks
