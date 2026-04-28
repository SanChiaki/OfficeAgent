# Ribbon Sync Workbook Change Log Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a visible `xISDP_Log` worksheet that records the latest 2000 Ribbon Sync upload/download cell changes.

**Architecture:** Keep the feature in `OfficeAgent.ExcelAddIn`: synchronization code knows resolved rows/columns and can produce audit entries, while a focused log store owns the workbook sheet layout and retention. Upload old values come from a pending edit tracker populated by Excel selection/change events; download old values are read immediately before overwrite.

**Tech Stack:** C# .NET Framework 4.8, Excel VSTO interop, xUnit, existing reflection-based ExcelAddIn tests.

---

## File Structure

- Create `src/OfficeAgent.ExcelAddIn/Excel/WorksheetChangeLogEntry.cs`
  - Internal DTO for one row in `xISDP_Log`.
- Create `src/OfficeAgent.ExcelAddIn/Excel/IWorksheetChangeLogStore.cs`
  - Internal append-only interface for sync execution.
- Create `src/OfficeAgent.ExcelAddIn/Excel/WorksheetChangeLogStore.cs`
  - Ensures `xISDP_Log`, rewrites headers, appends entries, trims to 2000.
- Create `src/OfficeAgent.ExcelAddIn/Excel/WorksheetCellAddress.cs`
  - Internal row/column coordinate DTO.
- Create `src/OfficeAgent.ExcelAddIn/Excel/WorksheetCellValue.cs`
  - Internal row/column/text DTO for before-edit snapshots.
- Create `src/OfficeAgent.ExcelAddIn/Excel/WorksheetPendingEditTracker.cs`
  - Stores first known before-edit Excel value until successful upload clears it.
- Modify `src/OfficeAgent.ExcelAddIn/Excel/IWorksheetGridAdapter.cs`
  - Add `EnsureWorksheetExists(string sheetName)`.
- Modify `src/OfficeAgent.ExcelAddIn/Excel/ExcelWorksheetGridAdapter.cs`
  - Implement `EnsureWorksheetExists` while preserving active worksheet.
- Modify `src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs`
  - Generate download/upload log entries after successful sync, skip unchanged/ID/missing ID cells, swallow log write failures into `OfficeAgentLog`.
- Modify `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs`
  - Instantiate log store and pending edit tracker, wire selection/change event capture.
- Modify `src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj`
  - Include new files.
- Modify `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs`
  - Add reflection-backed fake log store and tests for upload/download logging.
- Create `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetChangeLogStoreTests.cs`
  - Test retention and sheet layout through fake `IWorksheetGridAdapter`.
- Create `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetPendingEditTrackerTests.cs`
  - Test before-edit capture and successful-upload clearing.
- Modify `tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs`
  - Source-check ThisAddIn event wiring for pending edit tracking.
- Modify `docs/modules/ribbon-sync-current-behavior.md`
  - Document `xISDP_Log`.
- Modify `docs/vsto-manual-test-checklist.md`
  - Add manual validation items.

---

### Task 1: Add Log Store Contract And Retention Tests

**Files:**
- Create: `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetChangeLogStoreTests.cs`
- Create: `src/OfficeAgent.ExcelAddIn/Excel/WorksheetChangeLogEntry.cs`
- Create: `src/OfficeAgent.ExcelAddIn/Excel/IWorksheetChangeLogStore.cs`
- Create: `src/OfficeAgent.ExcelAddIn/Excel/WorksheetChangeLogStore.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/Excel/IWorksheetGridAdapter.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/Excel/ExcelWorksheetGridAdapter.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj`

- [ ] **Step 1: Write the failing log store test**

Create `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetChangeLogStoreTests.cs` with a reflection-backed fake grid. The first test creates 2001 entries, appends them, and asserts `xISDP_Log` contains one header row plus the newest 2000 rows, with `row-0002` as the first retained key and `row-2001` as the last key.

- [ ] **Step 2: Run the test and verify RED**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~WorksheetChangeLogStoreTests.AppendCreatesLogSheetAndKeepsLatestTwoThousandRows
```

Expected: FAIL because `WorksheetChangeLogStore` and related types do not exist.

- [ ] **Step 3: Implement minimal log store**

Add the DTO/interface/store files. `WorksheetChangeLogStore.Append` should:

```csharp
public void Append(IReadOnlyList<WorksheetChangeLogEntry> entries)
{
    var incoming = (entries ?? Array.Empty<WorksheetChangeLogEntry>())
        .Where(entry => entry != null)
        .ToArray();
    if (incoming.Length == 0)
    {
        return;
    }

    gridAdapter.EnsureWorksheetExists(LogSheetName);
    var existing = ReadExistingRows();
    var combined = existing.Concat(incoming).ToArray();
    var rows = combined.Skip(Math.Max(0, combined.Length - MaxEntries)).ToArray();
    RewriteRows(rows);
}
```

Implement `EnsureWorksheetExists` on the real grid adapter by finding an existing worksheet case-insensitively or adding one after the last worksheet.

- [ ] **Step 4: Run the log store test and verify GREEN**

Run the same `dotnet test ...WorksheetChangeLogStoreTests.AppendCreatesLogSheetAndKeepsLatestTwoThousandRows` command.

Expected: PASS.

---

### Task 2: Add Pending Edit Tracker

**Files:**
- Create: `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetPendingEditTrackerTests.cs`
- Create: `src/OfficeAgent.ExcelAddIn/Excel/WorksheetCellAddress.cs`
- Create: `src/OfficeAgent.ExcelAddIn/Excel/WorksheetCellValue.cs`
- Create: `src/OfficeAgent.ExcelAddIn/Excel/WorksheetPendingEditTracker.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj`

- [ ] **Step 1: Write failing tracker tests**

Create tests that verify:

```csharp
CaptureBeforeValues("Sheet1", cell(row: 6, column: 4, text: "2026-01-05"));
MarkChanged("Sheet1", address(row: 6, column: 4));
Assert.True(TryGetOriginalValue("Sheet1", 6, 4, out value));
Assert.Equal("2026-01-05", value);
Clear("Sheet1", 6, 4);
Assert.False(TryGetOriginalValue("Sheet1", 6, 4, out _));
```

Also verify a second edit keeps the first captured original value until clear.

- [ ] **Step 2: Run tests and verify RED**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~WorksheetPendingEditTrackerTests
```

Expected: FAIL because the tracker types do not exist.

- [ ] **Step 3: Implement minimal tracker**

Use two dictionaries:

```csharp
private readonly Dictionary<string, string> beforeValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
private readonly Dictionary<string, string> pendingOriginalValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
```

`CaptureBeforeValues` stores `sheet|row|column -> text`. `MarkChanged` adds a pending original only if one is not already pending and a before value exists. `Clear` removes both dictionaries for the coordinate.

- [ ] **Step 4: Run tracker tests and verify GREEN**

Run the same `dotnet test ...WorksheetPendingEditTrackerTests` command.

Expected: PASS.

---

### Task 3: Log Download Changes From Sync Execution

**Files:**
- Modify: `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs`

- [ ] **Step 1: Write failing partial download log test**

Add a test that executes a partial download over cell `(6,3)` where the old Excel value is `旧开始时间` and the downloaded value is `2026-02-01`. Assert fake log store receives one entry:

```csharp
Assert.Equal("row-1", entry.Key);
Assert.Equal("测试活动111/开始时间", entry.HeaderText);
Assert.Equal("下载", entry.ChangeMode);
Assert.Equal("2026-02-01", entry.NewValue);
Assert.Equal("旧开始时间", entry.OldValue);
```

- [ ] **Step 2: Run test and verify RED**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~WorksheetSyncExecutionServiceTests.ExecutePartialDownloadAppendsWorkbookLogForChangedCells
```

Expected: FAIL because `WorksheetSyncExecutionService` does not append log entries.

- [ ] **Step 3: Implement download log entry generation**

Add a constructor overload that accepts `IWorksheetChangeLogStore` and `WorksheetPendingEditTracker`, preserving the existing five-argument constructor. During partial/full download, read `OldValue` before writing, compute header text from `WorksheetColumnBinding`, skip ID and unchanged values, and call:

```csharp
AppendChangeLogEntries(entries);
```

`AppendChangeLogEntries` catches exceptions and calls:

```csharp
OfficeAgentLog.Error("ribbon-sync", "change-log.append.failed", "Failed to append workbook change log entries.", exception);
```

- [ ] **Step 4: Run test and verify GREEN**

Run the same partial download log test.

Expected: PASS.

---

### Task 4: Log Upload Changes After Successful BatchSave

**Files:**
- Modify: `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs`

- [ ] **Step 1: Write failing upload success test**

Seed pending original value for `(Sheet1, 6, 4)` as `2026-01-05`, upload current value `2026-01-10`, execute upload, and assert one log entry:

```csharp
Assert.Equal("row-1", entry.Key);
Assert.Equal("测试活动111/结束时间", entry.HeaderText);
Assert.Equal("上传", entry.ChangeMode);
Assert.Equal("2026-01-10", entry.NewValue);
Assert.Equal("2026-01-05", entry.OldValue);
```

- [ ] **Step 2: Write failing upload failure test**

Configure fake connector `BatchSave` to throw. Assert `ExecuteUpload` throws, fake log store remains empty, and `WorksheetPendingEditTracker.TryGetOriginalValue("Sheet1", 6, 4, out value)` still returns `2026-01-05`.

- [ ] **Step 3: Run upload tests and verify RED**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~WorksheetSyncExecutionServiceTests.ExecutePartialUploadAppendsWorkbookLogAfterSuccessfulBatchSave|FullyQualifiedName~WorksheetSyncExecutionServiceTests.ExecutePartialUploadDoesNotAppendLogOrClearPendingValueWhenBatchSaveFails"
```

Expected: FAIL because upload log candidates are not tracked.

- [ ] **Step 4: Implement upload log candidates**

Extend upload plan internals with log candidates and uploaded coordinates. `ExecuteUpload` should call `worksheetSyncService.Upload` first, then append upload log entries, then clear pending coordinates. If upload throws, do not append and do not clear.

- [ ] **Step 5: Run upload tests and verify GREEN**

Run the same filtered upload test command.

Expected: PASS.

---

### Task 5: Wire Excel Events And Runtime Construction

**Files:**
- Modify: `tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs`

- [ ] **Step 1: Write failing source-wiring test**

Add a source test asserting `ThisAddIn.cs` contains:

```csharp
internal WorksheetPendingEditTracker WorksheetPendingEditTracker { get; private set; }
internal IWorksheetChangeLogStore WorksheetChangeLogStore { get; private set; }
WorksheetPendingEditTracker.CaptureBeforeValues(sheetName, ReadWorksheetCellValues(target));
WorksheetPendingEditTracker.MarkChanged(sheetName, ReadWorksheetCellAddresses(target));
```

- [ ] **Step 2: Run test and verify RED**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~AgentRibbonConfigurationTests.ThisAddInTracksPendingCellEditsForWorkbookChangeLog
```

Expected: FAIL because wiring does not exist.

- [ ] **Step 3: Implement event wiring**

Create one `ExcelWorksheetGridAdapter` instance in startup, pass it to `WorksheetChangeLogStore` and `WorksheetSyncExecutionService`, instantiate `WorksheetPendingEditTracker`, and add event helper methods that skip `ISDP_Setting` and `xISDP_Log`.

- [ ] **Step 4: Run source-wiring test and verify GREEN**

Run the same AgentRibbonConfigurationTests filtered command.

Expected: PASS.

---

### Task 6: Update Documentation

**Files:**
- Modify: `docs/modules/ribbon-sync-current-behavior.md`
- Modify: `docs/vsto-manual-test-checklist.md`

- [ ] **Step 1: Update module behavior snapshot**

Add a concise `xISDP_Log` subsection documenting fixed columns, upload/download semantics, 2000-row retention, and exclusions.

- [ ] **Step 2: Update manual checklist**

Add checks for partial download logging, partial upload logging, and retention behavior.

- [ ] **Step 3: Review docs diff**

Run:

```powershell
git diff -- docs/modules/ribbon-sync-current-behavior.md docs/vsto-manual-test-checklist.md
```

Expected: diff only documents workbook sync change log behavior.

---

### Task 7: Full Verification

**Files:**
- All touched files.

- [ ] **Step 1: Run focused ExcelAddIn tests**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~WorksheetChangeLogStoreTests|FullyQualifiedName~WorksheetPendingEditTrackerTests|FullyQualifiedName~WorksheetSyncExecutionServiceTests|FullyQualifiedName~AgentRibbonConfigurationTests"
```

Expected: PASS.

- [ ] **Step 2: Run full ExcelAddIn test project**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj
```

Expected: PASS.

- [ ] **Step 3: Inspect final status**

Run:

```powershell
git status --short
```

Expected: only intended feature files and pre-existing unrelated user changes are present.
