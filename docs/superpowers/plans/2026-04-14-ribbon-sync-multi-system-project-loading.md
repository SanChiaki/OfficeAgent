# Ribbon Sync Multi-System Project Loading Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make Ribbon Sync load project options from registered business-system connectors and route later download/upload work by `systemKey`.

**Architecture:** Introduce a connector registry in `OfficeAgent.Core` as the single aggregation and lookup point. Refactor sync services and ribbon controller to depend on that registry, then adapt the current business connector and mock server so project options come from `/projects` instead of hardcoded data.

**Tech Stack:** C#, xUnit, VSTO Excel add-in, Node.js mock server

---

### Task 1: Lock Registry and Routing Behavior with Tests

**Files:**
- Modify: `tests/OfficeAgent.Core.Tests/SystemConnectorRegistryTests.cs`
- Modify: `tests/OfficeAgent.Core.Tests/WorksheetSyncServiceTests.cs`
- Modify: `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs`
- Create: `tests/OfficeAgent.ExcelAddIn.Tests/ProjectSelectionKeyTests.cs`

- [ ] **Step 1: Extend failing tests for registry normalization and execution re-routing**

Add coverage for:
- connector registry normalizing blank `ProjectOption.SystemKey`
- execution service reinitializing when stored binding `SystemKey` differs from selected project
- project selection key encoding/decoding preserving `systemKey + projectId`

- [ ] **Step 2: Run targeted tests to verify they fail for the intended reason**

Run:
```powershell
dotnet test tests/OfficeAgent.Core.Tests/OfficeAgent.Core.Tests.csproj --filter "FullyQualifiedName~SystemConnectorRegistryTests|FullyQualifiedName~WorksheetSyncServiceTests"
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~WorksheetSyncExecutionServiceTests|FullyQualifiedName~ProjectSelectionKeyTests"
```

Expected:
- compile/runtime failures because `SystemConnectorRegistry`, `ISystemConnector.SystemKey`, or project selection key helpers are not implemented yet

- [ ] **Step 3: Commit the red tests**

```bash
git add tests/OfficeAgent.Core.Tests/SystemConnectorRegistryTests.cs tests/OfficeAgent.Core.Tests/WorksheetSyncServiceTests.cs tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs tests/OfficeAgent.ExcelAddIn.Tests/ProjectSelectionKeyTests.cs
git commit -m "test: cover multi-system project routing"
```

### Task 2: Implement Connector Registry and Sync-Service Routing

**Files:**
- Create: `src/OfficeAgent.Core/Services/ISystemConnectorRegistry.cs`
- Create: `src/OfficeAgent.Core/Services/SystemConnectorRegistry.cs`
- Modify: `src/OfficeAgent.Core/Services/ISystemConnector.cs`
- Modify: `src/OfficeAgent.Core/Sync/WorksheetSyncService.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs`

- [ ] **Step 1: Add registry abstractions and `ISystemConnector.SystemKey`**

Implement a registry that:
- stores connectors by `SystemKey`
- aggregates all projects
- fills blank project `SystemKey` values from the connector
- throws clear exceptions for unknown keys

- [ ] **Step 2: Refactor sync services to route by `systemKey`**

Update `WorksheetSyncService` so it exposes:
- `GetProjects()`
- `CreateBindingSeed(string sheetName, ProjectOption project)`
- `LoadFieldMappingDefinition(string systemKey, string projectId)`
- `LoadFieldMappings(string sheetName, string systemKey, string projectId)`
- `Download(string systemKey, string projectId, ...)`
- `Upload(string systemKey, string projectId, ...)`

Update `WorksheetSyncExecutionService` so every runtime operation reads `binding.SystemKey` and passes it through.

- [ ] **Step 3: Refactor ribbon controller and add-in composition root**

Make `RibbonSyncController` depend on `WorksheetSyncService` instead of a direct connector for project loading/binding seed creation.

Wire `ThisAddIn` like:
```csharp
var connectors = new ISystemConnector[]
{
    new CurrentBusinessSystemConnector(() => SettingsStore.Load(), cookieContainer: SharedCookies.Container),
};
var connectorRegistry = new SystemConnectorRegistry(connectors);
WorksheetSyncService = new WorksheetSyncService(connectorRegistry, WorksheetMetadataStore, new WorksheetChangeTracker(), new SyncOperationPreviewFactory());
RibbonSyncController = new RibbonSyncController(WorksheetMetadataStore, WorksheetSyncService, GetActiveWorksheetName, WorksheetSyncExecutionService);
```

- [ ] **Step 4: Run targeted tests to verify green**

Run:
```powershell
dotnet test tests/OfficeAgent.Core.Tests/OfficeAgent.Core.Tests.csproj --filter "FullyQualifiedName~SystemConnectorRegistryTests|FullyQualifiedName~WorksheetSyncServiceTests"
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~WorksheetSyncExecutionServiceTests"
```

Expected:
- PASS

- [ ] **Step 5: Commit service-layer routing**

```bash
git add src/OfficeAgent.Core/Services/ISystemConnector.cs src/OfficeAgent.Core/Services/ISystemConnectorRegistry.cs src/OfficeAgent.Core/Services/SystemConnectorRegistry.cs src/OfficeAgent.Core/Sync/WorksheetSyncService.cs src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs src/OfficeAgent.ExcelAddIn/ThisAddIn.cs
git commit -m "feat: route worksheet sync by system key"
```

### Task 3: Remove Dropdown Collisions with Composite Project Keys

**Files:**
- Create: `src/OfficeAgent.ExcelAddIn/ProjectSelectionKey.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/AgentRibbon.cs`
- Modify: `tests/OfficeAgent.ExcelAddIn.Tests/RibbonSyncControllerTests.cs`
- Create: `tests/OfficeAgent.ExcelAddIn.Tests/ProjectSelectionKeyTests.cs`

- [ ] **Step 1: Add a small key helper**

Implement:
```csharp
internal static class ProjectSelectionKey
{
    public static string Build(string systemKey, string projectId) { ... }
    public static bool TryParse(string value, out string systemKey, out string projectId) { ... }
}
```

- [ ] **Step 2: Update ribbon dropdown state to use composite keys**

Replace `projectOptionsById` with a dictionary keyed by `ProjectSelectionKey.Build(project.SystemKey, project.ProjectId)`.

When syncing current selection from controller, build the same composite key from `ActiveSystemKey + ActiveProjectId`.

- [ ] **Step 3: Run targeted tests**

Run:
```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~ProjectSelectionKeyTests|FullyQualifiedName~RibbonSyncControllerTests"
```

Expected:
- PASS

- [ ] **Step 4: Commit UI keying**

```bash
git add src/OfficeAgent.ExcelAddIn/ProjectSelectionKey.cs src/OfficeAgent.ExcelAddIn/AgentRibbon.cs tests/OfficeAgent.ExcelAddIn.Tests/ProjectSelectionKeyTests.cs tests/OfficeAgent.ExcelAddIn.Tests/RibbonSyncControllerTests.cs
git commit -m "feat: key ribbon projects by system and project"
```

### Task 4: Fetch Project Options from the Current Business API

**Files:**
- Modify: `src/OfficeAgent.Infrastructure/Http/CurrentBusinessSystemConnector.cs`
- Modify: `tests/OfficeAgent.Infrastructure.Tests/CurrentBusinessSystemConnectorTests.cs`
- Modify: `tests/mock-server/server.js`
- Modify: `tests/mock-server/README.md`

- [ ] **Step 1: Write the failing API-backed project loading behavior**

Use the existing infrastructure test to require `/projects` calls and returned values.

- [ ] **Step 2: Implement `/projects` integration**

Make `CurrentBusinessSystemConnector.GetProjects()` call the business base URL `/projects` endpoint and map each result into `ProjectOption` while forcing `SystemKey = "current-business-system"` when missing.

Relax project validation so blank IDs are rejected, but dynamic API project IDs are allowed.

- [ ] **Step 3: Update mock server**

Add `/projects` response data aligned with the connector tests and current demo data.

- [ ] **Step 4: Run infrastructure and integration tests**

Run:
```powershell
dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj
dotnet test tests/OfficeAgent.IntegrationTests/OfficeAgent.IntegrationTests.csproj
```

Expected:
- PASS

- [ ] **Step 5: Commit API-backed project discovery**

```bash
git add src/OfficeAgent.Infrastructure/Http/CurrentBusinessSystemConnector.cs tests/OfficeAgent.Infrastructure.Tests/CurrentBusinessSystemConnectorTests.cs tests/mock-server/server.js tests/mock-server/README.md
git commit -m "feat: load ribbon projects from business api"
```

### Task 5: Refresh Product Documentation and Final Verification

**Files:**
- Modify: `docs/modules/ribbon-sync-current-behavior.md`
- Modify: `docs/ribbon-sync-real-system-integration-guide.md`
- Modify: `docs/module-index.md`
- Modify: `AGENTS.md`

- [ ] **Step 1: Document the new project-source architecture**

Explain:
- project list comes from registered connectors
- each binding persists `systemKey`
- future systems implement `ISystemConnector.GetProjects()`

- [ ] **Step 2: Run the full test matrix**

Run:
```powershell
dotnet test tests/OfficeAgent.Core.Tests/OfficeAgent.Core.Tests.csproj
dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj
dotnet test tests/OfficeAgent.IntegrationTests/OfficeAgent.IntegrationTests.csproj
```

Expected:
- PASS across all four suites

- [ ] **Step 3: Commit docs and verification**

```bash
git add docs/modules/ribbon-sync-current-behavior.md docs/ribbon-sync-real-system-integration-guide.md docs/module-index.md AGENTS.md
git commit -m "docs: describe multi-system ribbon project loading"
```
