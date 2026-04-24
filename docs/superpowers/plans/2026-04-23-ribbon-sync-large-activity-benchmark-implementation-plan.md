# Ribbon Sync Large Activity Benchmark Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a large mock Ribbon Sync project with 10,000 rows and two activity groups so performance testing covers dual-row headers plus large data writes.

**Architecture:** Keep the existing mock-server in-memory model and extend it with one generated benchmark project. Lock the behavior first with an integration test that expects the new project, row count, and activity-derived schema shape, then implement the generator and refresh the mock-server README.

**Tech Stack:** Node.js, Express, C#, xUnit

---

## File Structure

- Modify: `tests/OfficeAgent.IntegrationTests/CurrentBusinessSystemConnectorIntegrationTests.cs`
  Responsibility: assert the new benchmark project appears in `/projects`, returns 10,000 rows, and exposes activity-derived schema columns.
- Modify: `tests/mock-server/server.js`
  Responsibility: generate and register the large benchmark project while preserving current mock-server contracts.
- Modify: `tests/mock-server/README.md`
  Responsibility: document the new built-in benchmark project and its dataset shape.

### Task 1: Lock the benchmark contract with a failing integration test

**Files:**
- Modify: `tests/OfficeAgent.IntegrationTests/CurrentBusinessSystemConnectorIntegrationTests.cs`

- [ ] **Step 1: Write the failing test**

Add a new assertion slice to `GetProjectsReturnsAdditionalMockProjectsWithIndependentRows` or split it into a dedicated test:

```csharp
[Fact]
public async Task GetProjectsIncludesLargeBenchmarkProjectWithActivitySchema()
{
    var connector = await CreateConnectorAsync();

    var projects = connector.GetProjects();

    Assert.Equal(4, projects.Count);
    Assert.Contains(
        projects,
        project => string.Equals(project.ProjectId, "large-activity-benchmark", StringComparison.Ordinal)
            && string.Equals(project.DisplayName, "大数据活动压测项目", StringComparison.Ordinal));

    var rows = connector.Find("large-activity-benchmark", Array.Empty<string>(), Array.Empty<string>());
    Assert.Equal(10000, rows.Count);

    var firstRow = rows[0];
    Assert.Equal(11, firstRow.Count);
    Assert.Equal("benchmark-row-00001", firstRow["row_id"]?.ToString());
    Assert.True(firstRow.ContainsKey("name_benchmarka"));
    Assert.True(firstRow.ContainsKey("start_benchmarka"));
    Assert.True(firstRow.ContainsKey("end_benchmarka"));
    Assert.True(firstRow.ContainsKey("name_benchmarkb"));
    Assert.True(firstRow.ContainsKey("start_benchmarkb"));
    Assert.True(firstRow.ContainsKey("end_benchmarkb"));

    var schema = connector.GetSchema("large-activity-benchmark");
    Assert.Contains(schema.Columns, column => column.ApiFieldKey == "owner_name" && column.ColumnKind == WorksheetColumnKind.Single);
    Assert.Contains(schema.Columns, column => column.ApiFieldKey == "region" && column.ColumnKind == WorksheetColumnKind.Single);
    Assert.Contains(schema.Columns, column => column.ApiFieldKey == "priority" && column.ColumnKind == WorksheetColumnKind.Single);
    Assert.Contains(schema.Columns, column => column.ApiFieldKey == "status" && column.ColumnKind == WorksheetColumnKind.Single);
    Assert.Equal(6, schema.Columns.Count(column => column.ColumnKind == WorksheetColumnKind.ActivityProperty));
}
```

- [ ] **Step 2: Run the test to verify it fails**

Run:

```powershell
dotnet test tests/OfficeAgent.IntegrationTests/OfficeAgent.IntegrationTests.csproj --filter FullyQualifiedName~GetProjectsIncludesLargeBenchmarkProjectWithActivitySchema
```

Expected: FAIL because the mock server still exposes only the original three projects.

- [ ] **Step 3: Commit the red test checkpoint**

```powershell
git add tests/OfficeAgent.IntegrationTests/CurrentBusinessSystemConnectorIntegrationTests.cs
git commit -m "test: cover large activity benchmark mock project"
```

### Task 2: Implement the generated benchmark project in the mock server

**Files:**
- Modify: `tests/mock-server/server.js`

- [ ] **Step 1: Add the project generator**

Introduce helper functions in `tests/mock-server/server.js`:

```javascript
function createLargeActivityBenchmarkProject() {
  return {
    projectId: 'large-activity-benchmark',
    displayName: '大数据活动压测项目',
    headList: createLargeActivityBenchmarkHeadList(),
    rows: createLargeActivityBenchmarkRows(10000),
  };
}

function createLargeActivityBenchmarkHeadList() {
  return [
    { fieldKey: 'row_id', headerText: 'ID', headType: 'single', isId: true },
    { fieldKey: 'owner_name', headerText: '负责人', headType: 'single' },
    { fieldKey: 'region', headerText: '区域', headType: 'single' },
    { fieldKey: 'priority', headerText: '优先级', headType: 'single' },
    { fieldKey: 'status', headerText: '状态', headType: 'single' },
    { headType: 'activity', activityId: 'benchmarka', activityName: '基准阶段A' },
    { headType: 'activity', activityId: 'benchmarkb', activityName: '基准阶段B' },
  ];
}
```

- [ ] **Step 2: Generate 10,000 stable rows**

Add a deterministic row generator:

```javascript
function createLargeActivityBenchmarkRows(rowCount) {
  var regions = ['华东', '华北', '华南', '西南', '海外'];
  var priorities = ['P0', 'P1', 'P2', 'P3'];
  var statuses = ['未开始', '进行中', '已完成', '已暂停'];
  var rows = [];

  for (var index = 0; index < rowCount; index++) {
    var rowNumber = index + 1;
    var dayOffset = index % 28;
    rows.push({
      row_id: 'benchmark-row-' + String(rowNumber).padStart(5, '0'),
      owner_name: '区域负责人' + String((index % 300) + 1).padStart(3, '0'),
      region: regions[index % regions.length],
      priority: priorities[index % priorities.length],
      status: statuses[index % statuses.length],
      name_benchmarka: '阶段A-' + String(rowNumber).padStart(5, '0'),
      start_benchmarka: createIsoDate(2026, 1, 1 + dayOffset),
      end_benchmarka: createIsoDate(2026, 1, 3 + dayOffset),
      name_benchmarkb: '阶段B-' + String(rowNumber).padStart(5, '0'),
      start_benchmarkb: createIsoDate(2026, 1, 4 + dayOffset),
      end_benchmarkb: createIsoDate(2026, 1, 6 + dayOffset),
    });
  }

  return rows;
}
```

- [ ] **Step 3: Add small date helpers and register the project**

Register the project in `connectorProjectData`:

```javascript
"large-activity-benchmark": createLargeActivityBenchmarkProject(),
```

Add a safe helper so generated dates stay in `yyyy-MM-dd` format:

```javascript
function createIsoDate(year, month, day) {
  var date = new Date(Date.UTC(year, month - 1, day));
  return date.toISOString().slice(0, 10);
}
```

- [ ] **Step 4: Run the integration test to verify it passes**

Run:

```powershell
dotnet test tests/OfficeAgent.IntegrationTests/OfficeAgent.IntegrationTests.csproj --filter FullyQualifiedName~GetProjectsIncludesLargeBenchmarkProjectWithActivitySchema
```

Expected: PASS with the new project count, row count, and schema assertions satisfied.

- [ ] **Step 5: Commit the generator slice**

```powershell
git add tests/mock-server/server.js tests/OfficeAgent.IntegrationTests/CurrentBusinessSystemConnectorIntegrationTests.cs
git commit -m "test: add large activity benchmark mock project"
```

### Task 3: Update mock-server documentation and run final verification

**Files:**
- Modify: `tests/mock-server/README.md`
- Modify: `tests/OfficeAgent.IntegrationTests/CurrentBusinessSystemConnectorIntegrationTests.cs`
- Modify: `tests/mock-server/server.js`

- [ ] **Step 1: Update the README**

Add the benchmark project to the built-in project section:

```markdown
- `large-activity-benchmark`
  - `displayName = 大数据活动压测项目`
  - `10000` 条示例数据
  - `2` 个 activity 头：`benchmarka`、`benchmarkb`
  - 每行字段为 `row_id + 10` 个业务字段
```

- [ ] **Step 2: Run the focused integration suite**

Run:

```powershell
dotnet test tests/OfficeAgent.IntegrationTests/OfficeAgent.IntegrationTests.csproj --filter FullyQualifiedName~CurrentBusinessSystemConnectorIntegrationTests
```

Expected: PASS with `0` failed tests.

- [ ] **Step 3: Start the mock server once for a smoke check**

Run:

```powershell
node tests/mock-server/server.js
```

Expected: startup logs include `http://localhost:3200/projects` and the process stays alive without exceptions. Stop it manually after the smoke check.

- [ ] **Step 4: Commit the docs update**

```powershell
git add tests/mock-server/README.md
git commit -m "docs: describe large activity benchmark mock project"
```
