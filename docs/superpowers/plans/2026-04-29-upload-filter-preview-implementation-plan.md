# Upload Filter Preview Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add an optional Ribbon Sync upload filter that shows actual uploaded cells, skipped cells, and skip reasons before confirmation.

**Architecture:** Keep Excel selection parsing unchanged. Add a Core-level optional connector extension point, apply it during upload plan preparation, and keep `SyncOperationPreview.Changes` equal to the final submitted payload.

**Tech Stack:** C# net48, xUnit, VSTO add-in tests.

---

### Task 1: Preview Model And Factory

**Files:**
- Create: `src/OfficeAgent.Core/Models/SkippedCellChange.cs`
- Modify: `src/OfficeAgent.Core/Models/SyncOperationPreview.cs`
- Modify: `src/OfficeAgent.Core/Sync/SyncOperationPreviewFactory.cs`
- Test: `tests/OfficeAgent.Core.Tests/SyncOperationPreviewFactoryTests.cs`

- [ ] Write a failing test that passes included changes plus skipped changes with reasons to `CreateUploadPreview()`.
- [ ] Run `dotnet test tests/OfficeAgent.Core.Tests/OfficeAgent.Core.Tests.csproj --filter SyncOperationPreviewFactoryTests`.
- [ ] Add `SkippedCellChange`, add `SyncOperationPreview.SkippedChanges`, and update the factory summary/details.
- [ ] Re-run the same test command and confirm it passes.

### Task 2: Connector Filter Contract

**Files:**
- Create: `src/OfficeAgent.Core/Models/UploadChangeFilterResult.cs`
- Create: `src/OfficeAgent.Core/Services/IUploadChangeFilter.cs`
- Modify: `src/OfficeAgent.Core/Sync/WorksheetSyncService.cs`
- Test: `tests/OfficeAgent.Core.Tests/WorksheetSyncServiceTests.cs`

- [ ] Write a failing test proving a connector implementing `IUploadChangeFilter` returns included and skipped changes through `WorksheetSyncService`.
- [ ] Run `dotnet test tests/OfficeAgent.Core.Tests/OfficeAgent.Core.Tests.csproj --filter WorksheetSyncServiceTests`.
- [ ] Implement `FilterUploadChanges()` on `WorksheetSyncService`, defaulting to include-all when the connector does not implement the filter.
- [ ] Re-run the same test command and confirm it passes.

### Task 3: Excel Upload Plan Integration

**Files:**
- Modify: `src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs`
- Test: `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs`

- [ ] Write a failing test proving `PreparePartialUpload()` filters before preview and `ExecuteUpload()` submits only included changes.
- [ ] Run `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter WorksheetSyncExecutionServiceTests`.
- [ ] Apply filtering in `BuildUploadPreview()` using `WorksheetSyncService.FilterUploadChanges()`.
- [ ] Keep upload log candidates and pending uploaded cells aligned with `preview.Changes`, so skipped cells are not logged or cleared as uploaded.
- [ ] Re-run the same test command and confirm it passes.

### Task 4: Documentation And Verification

**Files:**
- Modify: `docs/modules/ribbon-sync-current-behavior.md`
- Modify: `docs/ribbon-sync-real-system-integration-guide.md`

- [ ] Update current behavior docs to describe upload filtering in the preview stage.
- [ ] Update real system integration guidance to describe `IUploadChangeFilter` and skip reasons.
- [ ] Run `dotnet test tests/OfficeAgent.Core.Tests/OfficeAgent.Core.Tests.csproj`.
- [ ] Run `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter WorksheetSyncExecutionServiceTests`.
