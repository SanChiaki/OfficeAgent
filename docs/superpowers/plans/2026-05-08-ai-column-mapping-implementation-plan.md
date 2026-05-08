# AI Column Mapping Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a standalone Ribbon action that uses the configured LLM to suggest mappings from the current sheet's real headers to initialized `SheetFieldMappings`, previews the result, and writes confirmed `Excel L1 / Excel L2` metadata.

**Architecture:** Keep the feature in the Ribbon Sync C# path. Core owns mapping contracts and deterministic validation/application; Infrastructure owns the OpenAI-compatible LLM call using existing settings; ExcelAddIn scans worksheet headers, hosts the preview dialog, and writes metadata only after confirmation.

**Tech Stack:** C#, .NET Framework 4.8, VSTO, WinForms, Newtonsoft.Json, xUnit

---

## File Structure

- `src/OfficeAgent.Core/Models/AiColumnMappingModels.cs`
  Responsibility: request, response, preview, and apply-result DTOs for AI column mapping.
- `src/OfficeAgent.Core/Services/IAiColumnMappingClient.cs`
  Responsibility: abstraction used by ExcelAddIn without depending on Infrastructure.
- `src/OfficeAgent.Core/Sync/AiColumnMappingService.cs`
  Responsibility: build LLM requests from metadata, validate LLM responses, create previews, and apply confirmed mappings to cloned `SheetFieldMappingRow` values.
- `tests/OfficeAgent.Core.Tests/AiColumnMappingServiceTests.cs`
  Responsibility: prove metadata identity fields are preserved and only eligible current header roles change.
- `src/OfficeAgent.Infrastructure/Http/AiColumnMappingClient.cs`
  Responsibility: call OpenAI-compatible chat completions using existing `AppSettings`.
- `tests/OfficeAgent.Infrastructure.Tests/AiColumnMappingClientTests.cs`
  Responsibility: verify HTTP payload, response parsing, and error handling.
- `src/OfficeAgent.ExcelAddIn/Excel/WorksheetHeaderScanner.cs`
  Responsibility: scan configured worksheet header rows into `AiColumnMappingActualHeader[]`.
- `src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs`
  Responsibility: orchestrate load metadata -> scan headers -> call AI client -> build preview -> save confirmed mapping rows.
- `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetHeaderScannerTests.cs`
  Responsibility: verify one-row and two-row complete header scanning.
- `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs`
  Responsibility: verify preview preparation and confirmed metadata writes.
- `src/OfficeAgent.ExcelAddIn/Dialogs/AiColumnMappingPreviewDialog.cs`
  Responsibility: WinForms preview confirmation UI.
- `src/OfficeAgent.ExcelAddIn/Dialogs/OperationResultDialog.cs`
  Responsibility: extend `IRibbonSyncDialogService` with AI mapping preview confirmation.
- `src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs`
  Responsibility: add the Ribbon command handler, precondition checks, confirmation, and localized result messages.
- `src/OfficeAgent.ExcelAddIn/AgentRibbon.Designer.cs`
  Responsibility: add the `AI map columns` Ribbon button under `Initialize sheet`.
- `src/OfficeAgent.ExcelAddIn/AgentRibbon.cs`
  Responsibility: wire the new button and localized label.
- `src/OfficeAgent.ExcelAddIn/Localization/HostLocalizedStrings.cs`
  Responsibility: add bilingual labels and messages.
- `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs`
  Responsibility: construct and inject `AiColumnMappingClient`.
- `tests/OfficeAgent.ExcelAddIn.Tests/RibbonSyncControllerTests.cs`
  Responsibility: verify controller confirmation flow.
- `tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs`
  Responsibility: verify Ribbon button placement and default label.
- `tests/OfficeAgent.ExcelAddIn.Tests/HostLocalizedStringsTests.cs`
  Responsibility: verify bilingual labels.
- `docs/modules/ribbon-sync-current-behavior.md`
  Responsibility: document the new user-visible Ribbon Sync behavior.
- `docs/ribbon-sync-real-system-integration-guide.md`
  Responsibility: explain that AI mapping updates `Excel L1 / Excel L2` only.
- `docs/vsto-manual-test-checklist.md`
  Responsibility: add manual validation steps.

## Implementation Notes

- Preserve the existing uncommitted `src/OfficeAgent.ExcelAddIn/Properties/Version.g.cs` change. Do not stage or revert it unless the user explicitly asks.
- Keep all business logic in C#. The React task pane is unchanged.
- Use `FieldMappingSemanticRole` and `RoleKey`; do not write code that depends on a fixed `SheetFieldMappings` display order.
- Treat accepted AI suggestions as suggestions, not truth. The deterministic Core service decides which rows can be written.
- Use a confidence threshold of `0.75`.
- If `HeaderType = single` and the accepted actual header has an `ActualL2`, apply it as grouped single when `HeaderRowCount > 1`; reject it when `HeaderRowCount = 1`.

### Task 1: Core AI Mapping Contracts and Application Service

**Files:**
- Create: `src/OfficeAgent.Core/Models/AiColumnMappingModels.cs`
- Create: `src/OfficeAgent.Core/Services/IAiColumnMappingClient.cs`
- Create: `src/OfficeAgent.Core/Sync/AiColumnMappingService.cs`
- Test: `tests/OfficeAgent.Core.Tests/AiColumnMappingServiceTests.cs`

- [ ] **Step 1: Write the failing Core tests**

Create `tests/OfficeAgent.Core.Tests/AiColumnMappingServiceTests.cs`:

```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Sync;
using Xunit;

namespace OfficeAgent.Core.Tests
{
    public sealed class AiColumnMappingServiceTests
    {
        [Fact]
        public void ApplyConfirmedPreviewUpdatesOnlyCurrentHeaderRoles()
        {
            var service = new AiColumnMappingService();
            var definition = BuildDefinition();
            var rows = BuildRows("Sheet1");
            var request = service.BuildRequest(
                "Sheet1",
                definition,
                rows,
                new[]
                {
                    new AiColumnMappingActualHeader { ExcelColumn = 2, ActualL1 = "项目负责人" },
                });
            var preview = service.CreatePreview(
                request,
                new AiColumnMappingResponse
                {
                    Mappings = new[]
                    {
                        new AiColumnMappingSuggestion
                        {
                            ExcelColumn = 2,
                            ActualL1 = "项目负责人",
                            TargetHeaderId = "owner_name",
                            TargetApiFieldKey = "owner_name",
                            Confidence = 0.93,
                            Reason = "same business meaning",
                        },
                    },
                },
                headerRowCount: 1);

            var result = service.ApplyConfirmedPreview("Sheet1", definition, rows, preview, headerRowCount: 1);
            var owner = result.Rows.Single(row => row.Values["HeaderId"] == "owner_name");

            Assert.Equal(1, result.AppliedCount);
            Assert.Equal("负责人", owner.Values["DefaultL1"]);
            Assert.Equal("项目负责人", owner.Values["CurrentL1"]);
            Assert.Equal(string.Empty, owner.Values["CurrentL2"]);
            Assert.Equal("owner_name", owner.Values["HeaderId"]);
            Assert.Equal("owner_name", owner.Values["ApiFieldKey"]);
            Assert.Equal("false", owner.Values["IsIdColumn"]);
        }

        [Fact]
        public void ApplyConfirmedPreviewLeavesLowConfidenceAndUnmatchedRowsUnchanged()
        {
            var service = new AiColumnMappingService();
            var definition = BuildDefinition();
            var rows = BuildRows("Sheet1");
            var request = service.BuildRequest(
                "Sheet1",
                definition,
                rows,
                new[]
                {
                    new AiColumnMappingActualHeader { ExcelColumn = 2, ActualL1 = "负责人简称" },
                });
            var preview = service.CreatePreview(
                request,
                new AiColumnMappingResponse
                {
                    Mappings = new[]
                    {
                        new AiColumnMappingSuggestion
                        {
                            ExcelColumn = 2,
                            ActualL1 = "负责人简称",
                            TargetHeaderId = "owner_name",
                            TargetApiFieldKey = "owner_name",
                            Confidence = 0.61,
                            Reason = "weak match",
                        },
                    },
                },
                headerRowCount: 1);

            var result = service.ApplyConfirmedPreview("Sheet1", definition, rows, preview, headerRowCount: 1);
            var owner = result.Rows.Single(row => row.Values["HeaderId"] == "owner_name");

            Assert.Equal(0, result.AppliedCount);
            Assert.Equal("负责人", owner.Values["CurrentL1"]);
            Assert.Equal(AiColumnMappingPreviewStatuses.LowConfidence, preview.Items.Single().Status);
        }

        [Fact]
        public void CreatePreviewRejectsDuplicateTargetFields()
        {
            var service = new AiColumnMappingService();
            var definition = BuildDefinition();
            var rows = BuildRows("Sheet1");
            var request = service.BuildRequest(
                "Sheet1",
                definition,
                rows,
                new[]
                {
                    new AiColumnMappingActualHeader { ExcelColumn = 2, ActualL1 = "负责人A" },
                    new AiColumnMappingActualHeader { ExcelColumn = 3, ActualL1 = "负责人B" },
                });

            var preview = service.CreatePreview(
                request,
                new AiColumnMappingResponse
                {
                    Mappings = new[]
                    {
                        new AiColumnMappingSuggestion { ExcelColumn = 2, TargetHeaderId = "owner_name", TargetApiFieldKey = "owner_name", Confidence = 0.9 },
                        new AiColumnMappingSuggestion { ExcelColumn = 3, TargetHeaderId = "owner_name", TargetApiFieldKey = "owner_name", Confidence = 0.9 },
                    },
                },
                headerRowCount: 1);

            Assert.Contains(preview.Items, item => item.ExcelColumn == 2 && item.Status == AiColumnMappingPreviewStatuses.Accepted);
            Assert.Contains(preview.Items, item => item.ExcelColumn == 3 && item.Status == AiColumnMappingPreviewStatuses.Rejected);
        }

        [Fact]
        public void ApplyConfirmedPreviewRejectsL2SuggestionsWhenHeaderRowCountIsOne()
        {
            var service = new AiColumnMappingService();
            var definition = BuildDefinition();
            var rows = BuildRows("Sheet1");
            var request = service.BuildRequest(
                "Sheet1",
                definition,
                rows,
                new[]
                {
                    new AiColumnMappingActualHeader { ExcelColumn = 2, ActualL1 = "基础信息", ActualL2 = "负责人" },
                });
            var preview = service.CreatePreview(
                request,
                new AiColumnMappingResponse
                {
                    Mappings = new[]
                    {
                        new AiColumnMappingSuggestion
                        {
                            ExcelColumn = 2,
                            ActualL1 = "基础信息",
                            ActualL2 = "负责人",
                            TargetHeaderId = "owner_name",
                            TargetApiFieldKey = "owner_name",
                            Confidence = 0.94,
                        },
                    },
                },
                headerRowCount: 1);

            var result = service.ApplyConfirmedPreview("Sheet1", definition, rows, preview, headerRowCount: 1);

            Assert.Equal(0, result.AppliedCount);
            Assert.Equal("负责人", result.Rows.Single(row => row.Values["HeaderId"] == "owner_name").Values["CurrentL1"]);
        }

        private static FieldMappingTableDefinition BuildDefinition()
        {
            return new FieldMappingTableDefinition
            {
                SystemKey = "current-business-system",
                Columns = new[]
                {
                    new FieldMappingColumnDefinition { ColumnName = "HeaderType", Role = FieldMappingSemanticRole.HeaderType },
                    new FieldMappingColumnDefinition { ColumnName = "ISDP L1", Role = FieldMappingSemanticRole.DefaultSingleHeaderText, RoleKey = "DefaultL1" },
                    new FieldMappingColumnDefinition { ColumnName = "ISDP L1", Role = FieldMappingSemanticRole.DefaultParentHeaderText, RoleKey = "DefaultL1" },
                    new FieldMappingColumnDefinition { ColumnName = "ISDP L2", Role = FieldMappingSemanticRole.DefaultChildHeaderText, RoleKey = "DefaultL2" },
                    new FieldMappingColumnDefinition { ColumnName = "Excel L1", Role = FieldMappingSemanticRole.CurrentSingleHeaderText, RoleKey = "CurrentL1" },
                    new FieldMappingColumnDefinition { ColumnName = "Excel L1", Role = FieldMappingSemanticRole.CurrentParentHeaderText, RoleKey = "CurrentL1" },
                    new FieldMappingColumnDefinition { ColumnName = "Excel L2", Role = FieldMappingSemanticRole.CurrentChildHeaderText, RoleKey = "CurrentL2" },
                    new FieldMappingColumnDefinition { ColumnName = "HeaderId", Role = FieldMappingSemanticRole.HeaderIdentity },
                    new FieldMappingColumnDefinition { ColumnName = "ApiFieldKey", Role = FieldMappingSemanticRole.ApiFieldKey },
                    new FieldMappingColumnDefinition { ColumnName = "IsIdColumn", Role = FieldMappingSemanticRole.IsIdColumn },
                    new FieldMappingColumnDefinition { ColumnName = "ActivityId", Role = FieldMappingSemanticRole.ActivityIdentity },
                    new FieldMappingColumnDefinition { ColumnName = "PropertyId", Role = FieldMappingSemanticRole.PropertyIdentity },
                },
            };
        }

        private static SheetFieldMappingRow[] BuildRows(string sheetName)
        {
            return new[]
            {
                CreateRow(sheetName, "single", "row_id", "ID", string.Empty, "true"),
                CreateRow(sheetName, "single", "owner_name", "负责人", string.Empty, "false"),
                CreateActivityRow(sheetName),
            };
        }

        private static SheetFieldMappingRow CreateRow(string sheetName, string headerType, string key, string l1, string l2, string isId)
        {
            return new SheetFieldMappingRow
            {
                SheetName = sheetName,
                Values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["HeaderType"] = headerType,
                    ["DefaultL1"] = l1,
                    ["DefaultL2"] = l2,
                    ["CurrentL1"] = l1,
                    ["CurrentL2"] = l2,
                    ["HeaderId"] = key,
                    ["ApiFieldKey"] = key,
                    ["IsIdColumn"] = isId,
                    ["ActivityId"] = string.Empty,
                    ["PropertyId"] = string.Empty,
                },
            };
        }

        private static SheetFieldMappingRow CreateActivityRow(string sheetName)
        {
            return new SheetFieldMappingRow
            {
                SheetName = sheetName,
                Values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["HeaderType"] = "activityProperty",
                    ["DefaultL1"] = "测试活动111",
                    ["DefaultL2"] = "开始时间",
                    ["CurrentL1"] = "测试活动111",
                    ["CurrentL2"] = "开始时间",
                    ["HeaderId"] = "start_12345678",
                    ["ApiFieldKey"] = "start_12345678",
                    ["IsIdColumn"] = "false",
                    ["ActivityId"] = "12345678",
                    ["PropertyId"] = "start",
                },
            };
        }
    }
}
```

- [ ] **Step 2: Run the Core tests and verify they fail**

Run: `dotnet test tests/OfficeAgent.Core.Tests/OfficeAgent.Core.Tests.csproj --filter FullyQualifiedName~AiColumnMappingServiceTests`

Expected: FAIL with missing type errors for `AiColumnMappingService`, `AiColumnMappingActualHeader`, and related DTOs.

- [ ] **Step 3: Add Core DTOs and client interface**

Create `src/OfficeAgent.Core/Models/AiColumnMappingModels.cs`:

```csharp
using System;

namespace OfficeAgent.Core.Models
{
    public static class AiColumnMappingPreviewStatuses
    {
        public const string Accepted = "accepted";
        public const string LowConfidence = "lowConfidence";
        public const string Unmatched = "unmatched";
        public const string Rejected = "rejected";
    }

    public sealed class AiColumnMappingActualHeader
    {
        public int ExcelColumn { get; set; }
        public string ActualL1 { get; set; } = string.Empty;
        public string ActualL2 { get; set; } = string.Empty;
        public string DisplayText { get; set; } = string.Empty;
    }

    public sealed class AiColumnMappingCandidate
    {
        public string HeaderId { get; set; } = string.Empty;
        public string HeaderType { get; set; } = string.Empty;
        public string ApiFieldKey { get; set; } = string.Empty;
        public string IsdpL1 { get; set; } = string.Empty;
        public string IsdpL2 { get; set; } = string.Empty;
        public string CurrentExcelL1 { get; set; } = string.Empty;
        public string CurrentExcelL2 { get; set; } = string.Empty;
        public bool IsIdColumn { get; set; }
    }

    public sealed class AiColumnMappingRequest
    {
        public string SheetName { get; set; } = string.Empty;
        public string SystemKey { get; set; } = string.Empty;
        public AiColumnMappingCandidate[] Candidates { get; set; } = Array.Empty<AiColumnMappingCandidate>();
        public AiColumnMappingActualHeader[] ActualHeaders { get; set; } = Array.Empty<AiColumnMappingActualHeader>();
    }

    public sealed class AiColumnMappingSuggestion
    {
        public int ExcelColumn { get; set; }
        public string ActualL1 { get; set; } = string.Empty;
        public string ActualL2 { get; set; } = string.Empty;
        public string TargetHeaderId { get; set; } = string.Empty;
        public string TargetApiFieldKey { get; set; } = string.Empty;
        public double Confidence { get; set; }
        public string Reason { get; set; } = string.Empty;
    }

    public sealed class AiColumnMappingUnmatchedHeader
    {
        public int ExcelColumn { get; set; }
        public string ActualL1 { get; set; } = string.Empty;
        public string ActualL2 { get; set; } = string.Empty;
        public string Reason { get; set; } = string.Empty;
    }

    public sealed class AiColumnMappingResponse
    {
        public AiColumnMappingSuggestion[] Mappings { get; set; } = Array.Empty<AiColumnMappingSuggestion>();
        public AiColumnMappingUnmatchedHeader[] Unmatched { get; set; } = Array.Empty<AiColumnMappingUnmatchedHeader>();
    }

    public sealed class AiColumnMappingPreview
    {
        public AiColumnMappingPreviewItem[] Items { get; set; } = Array.Empty<AiColumnMappingPreviewItem>();
    }

    public sealed class AiColumnMappingPreviewItem
    {
        public int ExcelColumn { get; set; }
        public string ActualL1 { get; set; } = string.Empty;
        public string ActualL2 { get; set; } = string.Empty;
        public string TargetHeaderId { get; set; } = string.Empty;
        public string TargetApiFieldKey { get; set; } = string.Empty;
        public string TargetIsdpL1 { get; set; } = string.Empty;
        public string TargetIsdpL2 { get; set; } = string.Empty;
        public string SuggestedExcelL1 { get; set; } = string.Empty;
        public string SuggestedExcelL2 { get; set; } = string.Empty;
        public double Confidence { get; set; }
        public string Reason { get; set; } = string.Empty;
        public string Status { get; set; } = AiColumnMappingPreviewStatuses.Unmatched;
    }

    public sealed class AiColumnMappingApplyResult
    {
        public SheetFieldMappingRow[] Rows { get; set; } = Array.Empty<SheetFieldMappingRow>();
        public int AppliedCount { get; set; }
        public int SkippedCount { get; set; }
    }
}
```

Create `src/OfficeAgent.Core/Services/IAiColumnMappingClient.cs`:

```csharp
using System.Threading.Tasks;
using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Services
{
    public interface IAiColumnMappingClient
    {
        AiColumnMappingResponse Map(AiColumnMappingRequest request);

        Task<AiColumnMappingResponse> MapAsync(AiColumnMappingRequest request);
    }
}
```

- [ ] **Step 4: Add the deterministic Core service**

Create `src/OfficeAgent.Core/Sync/AiColumnMappingService.cs` with the following implementation:

```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Sync
{
    public sealed class AiColumnMappingService
    {
        public const double DefaultConfidenceThreshold = 0.75;
        private readonly FieldMappingValueAccessor valueAccessor = new FieldMappingValueAccessor();

        public AiColumnMappingRequest BuildRequest(
            string sheetName,
            FieldMappingTableDefinition definition,
            IReadOnlyList<SheetFieldMappingRow> mappings,
            IReadOnlyList<AiColumnMappingActualHeader> actualHeaders)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                throw new ArgumentException("Sheet name is required.", nameof(sheetName));
            }

            if (definition == null)
            {
                throw new ArgumentNullException(nameof(definition));
            }

            var rows = (mappings ?? Array.Empty<SheetFieldMappingRow>())
                .Where(row => row != null && string.Equals(row.SheetName, sheetName, StringComparison.OrdinalIgnoreCase))
                .ToArray();

            return new AiColumnMappingRequest
            {
                SheetName = sheetName,
                SystemKey = definition.SystemKey ?? string.Empty,
                Candidates = rows.Select(row => new AiColumnMappingCandidate
                {
                    HeaderId = valueAccessor.GetValue(definition, row, FieldMappingSemanticRole.HeaderIdentity),
                    HeaderType = valueAccessor.GetValue(definition, row, FieldMappingSemanticRole.HeaderType),
                    ApiFieldKey = valueAccessor.GetValue(definition, row, FieldMappingSemanticRole.ApiFieldKey),
                    IsdpL1 = ResolveDefaultL1(definition, row),
                    IsdpL2 = valueAccessor.GetValue(definition, row, FieldMappingSemanticRole.DefaultChildHeaderText),
                    CurrentExcelL1 = ResolveCurrentL1(definition, row),
                    CurrentExcelL2 = valueAccessor.GetValue(definition, row, FieldMappingSemanticRole.CurrentChildHeaderText),
                    IsIdColumn = valueAccessor.GetBoolean(definition, row, FieldMappingSemanticRole.IsIdColumn),
                }).Where(candidate => !string.IsNullOrWhiteSpace(candidate.ApiFieldKey)).ToArray(),
                ActualHeaders = (actualHeaders ?? Array.Empty<AiColumnMappingActualHeader>())
                    .Where(header => header != null && header.ExcelColumn > 0)
                    .Select(header => new AiColumnMappingActualHeader
                    {
                        ExcelColumn = header.ExcelColumn,
                        ActualL1 = header.ActualL1 ?? string.Empty,
                        ActualL2 = header.ActualL2 ?? string.Empty,
                        DisplayText = string.IsNullOrWhiteSpace(header.DisplayText)
                            ? BuildDisplayText(header.ActualL1, header.ActualL2)
                            : header.DisplayText,
                    }).ToArray(),
            };
        }

        public AiColumnMappingPreview CreatePreview(
            AiColumnMappingRequest request,
            AiColumnMappingResponse response,
            int headerRowCount)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            var candidatesByHeaderId = request.Candidates.ToDictionary(item => item.HeaderId ?? string.Empty, StringComparer.OrdinalIgnoreCase);
            var candidatesByApiFieldKey = request.Candidates.ToDictionary(item => item.ApiFieldKey ?? string.Empty, StringComparer.OrdinalIgnoreCase);
            var actualHeadersByColumn = request.ActualHeaders.ToDictionary(item => item.ExcelColumn);
            var usedTargets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var usedColumns = new HashSet<int>();
            var items = new List<AiColumnMappingPreviewItem>();

            foreach (var suggestion in response?.Mappings ?? Array.Empty<AiColumnMappingSuggestion>())
            {
                if (suggestion == null || !actualHeadersByColumn.TryGetValue(suggestion.ExcelColumn, out var actual))
                {
                    continue;
                }

                usedColumns.Add(suggestion.ExcelColumn);
                var candidate = ResolveCandidate(candidatesByHeaderId, candidatesByApiFieldKey, suggestion);
                var status = ResolveStatus(candidate, suggestion, usedTargets, headerRowCount);
                if (candidate != null && string.Equals(status, AiColumnMappingPreviewStatuses.Accepted, StringComparison.Ordinal))
                {
                    usedTargets.Add(candidate.HeaderId);
                }

                items.Add(CreatePreviewItem(actual, candidate, suggestion, status));
            }

            foreach (var actual in request.ActualHeaders.Where(header => !usedColumns.Contains(header.ExcelColumn)))
            {
                items.Add(CreatePreviewItem(actual, null, null, AiColumnMappingPreviewStatuses.Unmatched));
            }

            return new AiColumnMappingPreview
            {
                Items = items.OrderBy(item => item.ExcelColumn).ToArray(),
            };
        }

        public AiColumnMappingApplyResult ApplyConfirmedPreview(
            string sheetName,
            FieldMappingTableDefinition definition,
            IReadOnlyList<SheetFieldMappingRow> mappings,
            AiColumnMappingPreview preview,
            int headerRowCount)
        {
            if (definition == null)
            {
                throw new ArgumentNullException(nameof(definition));
            }

            var acceptedByHeaderId = (preview?.Items ?? Array.Empty<AiColumnMappingPreviewItem>())
                .Where(item => item != null && string.Equals(item.Status, AiColumnMappingPreviewStatuses.Accepted, StringComparison.Ordinal))
                .Where(item => !(headerRowCount <= 1 && !string.IsNullOrWhiteSpace(item.SuggestedExcelL2)))
                .GroupBy(item => item.TargetHeaderId ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);

            var applied = 0;
            var rows = new List<SheetFieldMappingRow>();
            foreach (var row in mappings ?? Array.Empty<SheetFieldMappingRow>())
            {
                if (row == null || !string.Equals(row.SheetName, sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    rows.Add(CloneRow(row));
                    continue;
                }

                var headerId = valueAccessor.GetValue(definition, row, FieldMappingSemanticRole.HeaderIdentity);
                if (!acceptedByHeaderId.TryGetValue(headerId, out var accepted))
                {
                    rows.Add(CloneRow(row));
                    continue;
                }

                rows.Add(ApplyAccepted(definition, row, accepted, headerRowCount));
                applied++;
            }

            return new AiColumnMappingApplyResult
            {
                Rows = rows.ToArray(),
                AppliedCount = applied,
                SkippedCount = Math.Max(0, (preview?.Items?.Length ?? 0) - applied),
            };
        }

        private string ResolveDefaultL1(FieldMappingTableDefinition definition, SheetFieldMappingRow row)
        {
            var single = valueAccessor.GetValue(definition, row, FieldMappingSemanticRole.DefaultSingleHeaderText);
            return string.IsNullOrWhiteSpace(single)
                ? valueAccessor.GetValue(definition, row, FieldMappingSemanticRole.DefaultParentHeaderText)
                : single;
        }

        private string ResolveCurrentL1(FieldMappingTableDefinition definition, SheetFieldMappingRow row)
        {
            var single = valueAccessor.GetValue(definition, row, FieldMappingSemanticRole.CurrentSingleHeaderText);
            return string.IsNullOrWhiteSpace(single)
                ? valueAccessor.GetValue(definition, row, FieldMappingSemanticRole.CurrentParentHeaderText)
                : single;
        }

        private static AiColumnMappingCandidate ResolveCandidate(
            IReadOnlyDictionary<string, AiColumnMappingCandidate> candidatesByHeaderId,
            IReadOnlyDictionary<string, AiColumnMappingCandidate> candidatesByApiFieldKey,
            AiColumnMappingSuggestion suggestion)
        {
            if (suggestion == null)
            {
                return null;
            }

            if (!string.IsNullOrWhiteSpace(suggestion.TargetHeaderId) &&
                candidatesByHeaderId.TryGetValue(suggestion.TargetHeaderId, out var byHeaderId))
            {
                return byHeaderId;
            }

            return !string.IsNullOrWhiteSpace(suggestion.TargetApiFieldKey) &&
                   candidatesByApiFieldKey.TryGetValue(suggestion.TargetApiFieldKey, out var byApiFieldKey)
                ? byApiFieldKey
                : null;
        }

        private static string ResolveStatus(
            AiColumnMappingCandidate candidate,
            AiColumnMappingSuggestion suggestion,
            ISet<string> usedTargets,
            int headerRowCount)
        {
            if (candidate == null)
            {
                return AiColumnMappingPreviewStatuses.Rejected;
            }

            if (usedTargets.Contains(candidate.HeaderId))
            {
                return AiColumnMappingPreviewStatuses.Rejected;
            }

            if (headerRowCount <= 1 && !string.IsNullOrWhiteSpace(suggestion?.ActualL2))
            {
                return AiColumnMappingPreviewStatuses.Rejected;
            }

            return suggestion?.Confidence >= DefaultConfidenceThreshold
                ? AiColumnMappingPreviewStatuses.Accepted
                : AiColumnMappingPreviewStatuses.LowConfidence;
        }

        private static AiColumnMappingPreviewItem CreatePreviewItem(
            AiColumnMappingActualHeader actual,
            AiColumnMappingCandidate candidate,
            AiColumnMappingSuggestion suggestion,
            string status)
        {
            return new AiColumnMappingPreviewItem
            {
                ExcelColumn = actual?.ExcelColumn ?? suggestion?.ExcelColumn ?? 0,
                ActualL1 = actual?.ActualL1 ?? suggestion?.ActualL1 ?? string.Empty,
                ActualL2 = actual?.ActualL2 ?? suggestion?.ActualL2 ?? string.Empty,
                TargetHeaderId = candidate?.HeaderId ?? suggestion?.TargetHeaderId ?? string.Empty,
                TargetApiFieldKey = candidate?.ApiFieldKey ?? suggestion?.TargetApiFieldKey ?? string.Empty,
                TargetIsdpL1 = candidate?.IsdpL1 ?? string.Empty,
                TargetIsdpL2 = candidate?.IsdpL2 ?? string.Empty,
                SuggestedExcelL1 = actual?.ActualL1 ?? suggestion?.ActualL1 ?? string.Empty,
                SuggestedExcelL2 = actual?.ActualL2 ?? suggestion?.ActualL2 ?? string.Empty,
                Confidence = suggestion?.Confidence ?? 0,
                Reason = suggestion?.Reason ?? string.Empty,
                Status = status ?? AiColumnMappingPreviewStatuses.Unmatched,
            };
        }

        private static SheetFieldMappingRow ApplyAccepted(
            FieldMappingTableDefinition definition,
            SheetFieldMappingRow row,
            AiColumnMappingPreviewItem accepted,
            int headerRowCount)
        {
            var clone = CloneRow(row);
            var currentSingleKey = GetValueKey(definition, FieldMappingSemanticRole.CurrentSingleHeaderText);
            var currentParentKey = GetValueKey(definition, FieldMappingSemanticRole.CurrentParentHeaderText);
            var currentChildKey = GetValueKey(definition, FieldMappingSemanticRole.CurrentChildHeaderText);
            var values = new Dictionary<string, string>(clone.Values, StringComparer.OrdinalIgnoreCase);

            if (!string.IsNullOrWhiteSpace(currentSingleKey))
            {
                values[currentSingleKey] = accepted.SuggestedExcelL1 ?? string.Empty;
            }

            if (!string.IsNullOrWhiteSpace(currentParentKey))
            {
                values[currentParentKey] = accepted.SuggestedExcelL1 ?? string.Empty;
            }

            if (!string.IsNullOrWhiteSpace(currentChildKey))
            {
                values[currentChildKey] = headerRowCount > 1 ? accepted.SuggestedExcelL2 ?? string.Empty : string.Empty;
            }

            clone.Values = values;
            return clone;
        }

        private static string GetValueKey(FieldMappingTableDefinition definition, FieldMappingSemanticRole role)
        {
            return (definition?.Columns ?? Array.Empty<FieldMappingColumnDefinition>())
                .Where(column => column != null && column.Role == role)
                .Select(column => string.IsNullOrWhiteSpace(column.RoleKey) ? column.ColumnName : column.RoleKey)
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? string.Empty;
        }

        private static string BuildDisplayText(string l1, string l2)
        {
            return string.IsNullOrWhiteSpace(l2) ? l1 ?? string.Empty : (l1 ?? string.Empty) + "/" + l2;
        }

        private static SheetFieldMappingRow CloneRow(SheetFieldMappingRow row)
        {
            return new SheetFieldMappingRow
            {
                SheetName = row?.SheetName ?? string.Empty,
                Values = new Dictionary<string, string>(row?.Values ?? new Dictionary<string, string>(), StringComparer.OrdinalIgnoreCase),
            };
        }
    }
}
```

- [ ] **Step 5: Run the Core tests and verify they pass**

Run: `dotnet test tests/OfficeAgent.Core.Tests/OfficeAgent.Core.Tests.csproj --filter FullyQualifiedName~AiColumnMappingServiceTests`

Expected: PASS.

- [ ] **Step 6: Commit Task 1**

```powershell
git add -- src/OfficeAgent.Core/Models/AiColumnMappingModels.cs src/OfficeAgent.Core/Services/IAiColumnMappingClient.cs src/OfficeAgent.Core/Sync/AiColumnMappingService.cs tests/OfficeAgent.Core.Tests/AiColumnMappingServiceTests.cs
git commit -m "feat: add ai column mapping core"
```

### Task 2: OpenAI-Compatible AI Column Mapping Client

**Files:**
- Create: `src/OfficeAgent.Infrastructure/Http/AiColumnMappingClient.cs`
- Test: `tests/OfficeAgent.Infrastructure.Tests/AiColumnMappingClientTests.cs`

- [ ] **Step 1: Write the failing Infrastructure tests**

Create `tests/OfficeAgent.Infrastructure.Tests/AiColumnMappingClientTests.cs`:

```csharp
using System;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using OfficeAgent.Core.Models;
using OfficeAgent.Infrastructure.Http;
using Xunit;

namespace OfficeAgent.Infrastructure.Tests
{
    public sealed class AiColumnMappingClientTests
    {
        [Fact]
        public void MapPostsColumnMappingRequestToChatCompletions()
        {
            var handler = new RecordingHandler(_ => new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent("{\"choices\":[{\"message\":{\"content\":\"{\\\"mappings\\\":[{\\\"excelColumn\\\":2,\\\"targetHeaderId\\\":\\\"owner_name\\\",\\\"targetApiFieldKey\\\":\\\"owner_name\\\",\\\"confidence\\\":0.91,\\\"reason\\\":\\\"match\\\"}],\\\"unmatched\\\":[]}\"}}]}"),
            });
            var client = new AiColumnMappingClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BaseUrl = "https://api.internal.example",
                    Model = "gpt-5-mini",
                });

            var response = client.Map(new AiColumnMappingRequest
            {
                SheetName = "Sheet1",
                Candidates = new[] { new AiColumnMappingCandidate { HeaderId = "owner_name", ApiFieldKey = "owner_name", IsdpL1 = "负责人" } },
                ActualHeaders = new[] { new AiColumnMappingActualHeader { ExcelColumn = 2, ActualL1 = "项目负责人" } },
            });

            Assert.Equal("https://api.internal.example/v1/chat/completions", handler.LastRequest.RequestUri.ToString());
            Assert.Equal("Bearer", handler.LastRequest.Headers.Authorization?.Scheme);
            Assert.Equal("secret-token", handler.LastRequest.Headers.Authorization?.Parameter);
            Assert.Contains("gpt-5-mini", handler.LastBody, StringComparison.Ordinal);
            Assert.Contains("\"response_format\":{\"type\":\"json_object\"}", handler.LastBody, StringComparison.Ordinal);
            Assert.Contains("项目负责人", handler.LastBody, StringComparison.Ordinal);
            Assert.Equal("owner_name", response.Mappings[0].TargetHeaderId);
        }

        [Fact]
        public void MapRejectsMissingApiKeys()
        {
            var handler = new RecordingHandler(_ => new HttpResponseMessage(HttpStatusCode.OK));
            var client = new AiColumnMappingClient(new HttpClient(handler), () => new AppSettings { ApiKey = " ", BaseUrl = "https://api.internal.example" });

            var error = Assert.Throws<InvalidOperationException>(() => client.Map(new AiColumnMappingRequest()));

            Assert.Contains("API Key", error.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(0, handler.CallCount);
        }

        [Fact]
        public void MapRejectsNonJsonModelContent()
        {
            var handler = new RecordingHandler(_ => new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent("{\"choices\":[{\"message\":{\"content\":\"not-json\"}}]}"),
            });
            var client = new AiColumnMappingClient(new HttpClient(handler), () => new AppSettings { ApiKey = "secret-token", BaseUrl = "https://api.internal.example" });

            var error = Assert.Throws<InvalidOperationException>(() => client.Map(new AiColumnMappingRequest()));

            Assert.Contains("AI column mapping", error.Message, StringComparison.OrdinalIgnoreCase);
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

            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
            {
                CallCount++;
                LastRequest = request;
                LastBody = request.Content?.ReadAsStringAsync().GetAwaiter().GetResult() ?? string.Empty;
                return Task.FromResult(responder(request));
            }
        }
    }
}
```

- [ ] **Step 2: Run the Infrastructure tests and verify they fail**

Run: `dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj --filter FullyQualifiedName~AiColumnMappingClientTests`

Expected: FAIL with missing type `AiColumnMappingClient`.

- [ ] **Step 3: Add the LLM client**

Create `src/OfficeAgent.Infrastructure/Http/AiColumnMappingClient.cs`. Use the same base URL and chat completions behavior as `LlmPlannerClient`, but return `AiColumnMappingResponse`:

```csharp
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Authentication;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;

namespace OfficeAgent.Infrastructure.Http
{
    public sealed class AiColumnMappingClient : IAiColumnMappingClient
    {
        private readonly HttpClient httpClient;
        private readonly Func<AppSettings> loadSettings;

        public AiColumnMappingClient(Func<AppSettings> loadSettings)
            : this(null, loadSettings)
        {
        }

        public AiColumnMappingClient(HttpClient httpClient, Func<AppSettings> loadSettings)
        {
            this.httpClient = httpClient ?? new HttpClient(new HttpClientHandler
            {
                SslProtocols = SslProtocols.Tls12 | SslProtocols.Tls13,
            })
            {
                Timeout = TimeSpan.FromSeconds(120),
            };
            this.loadSettings = loadSettings ?? throw new ArgumentNullException(nameof(loadSettings));
        }

        public AiColumnMappingResponse Map(AiColumnMappingRequest request)
        {
            return MapAsync(request).GetAwaiter().GetResult();
        }

        public async Task<AiColumnMappingResponse> MapAsync(AiColumnMappingRequest request)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            var settings = loadSettings() ?? new AppSettings();
            if (string.IsNullOrWhiteSpace(settings.ApiKey))
            {
                throw new InvalidOperationException("An API Key is required before AI column mapping can call the model API.");
            }

            var baseUrl = AppSettings.NormalizeBaseUrl(settings.BaseUrl);
            if (!Uri.TryCreate(baseUrl, UriKind.Absolute, out var baseUri) ||
                (baseUri.Scheme != Uri.UriSchemeHttp && baseUri.Scheme != Uri.UriSchemeHttps))
            {
                throw new InvalidOperationException("The configured AI column mapping Base URL is invalid. Update settings and try again.");
            }

            var endpoint = BuildChatCompletionsEndpoint(baseUri);
            var payload = JsonConvert.SerializeObject(new
            {
                model = settings.Model,
                messages = new[]
                {
                    new { role = "system", content = BuildSystemPrompt() },
                    new { role = "user", content = JsonConvert.SerializeObject(request, Formatting.Indented) },
                },
                response_format = new { type = "json_object" },
            });

            using (var httpRequest = new HttpRequestMessage(HttpMethod.Post, endpoint))
            {
                httpRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", settings.ApiKey);
                httpRequest.Content = new StringContent(payload, Encoding.UTF8, "application/json");
                using (var response = await httpClient.SendAsync(httpRequest).ConfigureAwait(false))
                {
                    var body = await (response.Content?.ReadAsStringAsync() ?? Task.FromResult(string.Empty)).ConfigureAwait(false);
                    if (!response.IsSuccessStatusCode)
                    {
                        throw new InvalidOperationException($"AI column mapping request failed ({(int)response.StatusCode} {response.ReasonPhrase}): {body}");
                    }

                    return ParseResponse(ExtractChatText(body));
                }
            }
        }
    }
}
```

Complete the private helpers in the same file:

- `BuildChatCompletionsEndpoint(Uri baseUri)` identical to `LlmPlannerClient`.
- `BuildSystemPrompt()` includes the constraints from the design doc: JSON only, no invented fields, one actual header to one target field, preserve ID fields unless obvious.
- `ExtractChatText(string responseBody)` supports string `message.content` and text/output_text array items like `LlmPlannerClient`.
- `ParseResponse(string text)` calls `JsonConvert.DeserializeObject<AiColumnMappingResponse>()` and rejects null, malformed JSON, or missing `mappings`.

- [ ] **Step 4: Run the Infrastructure tests and verify they pass**

Run: `dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj --filter FullyQualifiedName~AiColumnMappingClientTests`

Expected: PASS.

- [ ] **Step 5: Commit Task 2**

```powershell
git add -- src/OfficeAgent.Infrastructure/Http/AiColumnMappingClient.cs tests/OfficeAgent.Infrastructure.Tests/AiColumnMappingClientTests.cs
git commit -m "feat: add ai column mapping client"
```

### Task 3: Header Scanning and Execution Service Integration

**Files:**
- Create: `src/OfficeAgent.ExcelAddIn/Excel/WorksheetHeaderScanner.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj`
- Modify: `src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs`
- Test: `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetHeaderScannerTests.cs`
- Test: `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs`

- [ ] **Step 1: Write the failing header scanner tests**

Create `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetHeaderScannerTests.cs`:

```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using OfficeAgent.Core.Models;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class WorksheetHeaderScannerTests
    {
        [Fact]
        public void ScanReadsCompleteSingleRowHeaderArea()
        {
            var scanner = CreateScanner();
            var grid = new FakeWorksheetGridAdapter(LoadGridInterfaceType());
            grid.SetCell("Sheet1", 3, 1, "ID");
            grid.SetCell("Sheet1", 3, 2, "项目负责人");
            grid.SetCell("Sheet1", 3, 3, "用户备注");

            var headers = InvokeScan(
                scanner,
                "Sheet1",
                new SheetBinding { SheetName = "Sheet1", HeaderStartRow = 3, HeaderRowCount = 1 },
                grid);

            Assert.Equal(new[] { 1, 2, 3 }, headers.Select(header => header.ExcelColumn));
            Assert.Equal("项目负责人", headers[1].ActualL1);
            Assert.Equal(string.Empty, headers[1].ActualL2);
        }

        [Fact]
        public void ScanCarriesForwardTwoRowParentHeaders()
        {
            var scanner = CreateScanner();
            var grid = new FakeWorksheetGridAdapter(LoadGridInterfaceType());
            grid.SetCell("Sheet1", 3, 1, "ID");
            grid.SetCell("Sheet1", 3, 2, "测试活动111");
            grid.SetCell("Sheet1", 4, 2, "开始时间");
            grid.SetCell("Sheet1", 4, 3, "结束时间");

            var headers = InvokeScan(
                scanner,
                "Sheet1",
                new SheetBinding { SheetName = "Sheet1", HeaderStartRow = 3, HeaderRowCount = 2 },
                grid);

            Assert.Equal("测试活动111", headers.Single(header => header.ExcelColumn == 3).ActualL1);
            Assert.Equal("结束时间", headers.Single(header => header.ExcelColumn == 3).ActualL2);
        }

        private static object CreateScanner()
        {
            var assembly = LoadAddInAssembly();
            var type = assembly.GetType("OfficeAgent.ExcelAddIn.Excel.WorksheetHeaderScanner", throwOnError: true);
            return Activator.CreateInstance(type);
        }

        private static AiColumnMappingActualHeader[] InvokeScan(
            object scanner,
            string sheetName,
            SheetBinding binding,
            FakeWorksheetGridAdapter grid)
        {
            var method = scanner.GetType().GetMethod("Scan", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            return (AiColumnMappingActualHeader[])method.Invoke(scanner, new[] { sheetName, binding, grid.GetTransparentProxy() });
        }

        private static Type LoadGridInterfaceType()
        {
            return LoadAddInAssembly().GetType("OfficeAgent.ExcelAddIn.Excel.IWorksheetGridAdapter", throwOnError: true);
        }

        private static Assembly LoadAddInAssembly()
        {
            return Assembly.LoadFrom(ResolveRepositoryPath("src", "OfficeAgent.ExcelAddIn", "bin", "Debug", "OfficeAgent.ExcelAddIn.dll"));
        }

        private static string ResolveRepositoryPath(params string[] segments)
        {
            return System.IO.Path.GetFullPath(System.IO.Path.Combine(new[]
            {
                AppContext.BaseDirectory,
                "..",
                "..",
                "..",
                "..",
                "..",
            }.Concat(segments).ToArray()));
        }

        private sealed class FakeWorksheetGridAdapter : RealProxy
        {
            private readonly Dictionary<string, string> cells = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            public FakeWorksheetGridAdapter(Type interfaceType)
                : base(interfaceType)
            {
            }

            public void SetCell(string sheetName, int row, int column, string value)
            {
                cells[$"{sheetName}!{row}!{column}"] = value ?? string.Empty;
            }

            public override IMessage Invoke(IMessage msg)
            {
                var call = (IMethodCallMessage)msg;
                switch (call.MethodName)
                {
                    case "GetCellText":
                        return new ReturnMessage(
                            GetCell((string)call.InArgs[0], (int)call.InArgs[1], (int)call.InArgs[2]),
                            null,
                            0,
                            call.LogicalCallContext,
                            call);
                    case "GetLastUsedColumn":
                        return new ReturnMessage(GetLastUsedColumn((string)call.InArgs[0]), null, 0, call.LogicalCallContext, call);
                    case "BeginBulkOperation":
                        return new ReturnMessage(new NoopDisposable(), null, 0, call.LogicalCallContext, call);
                    default:
                        return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
                }
            }

            public new object GetTransparentProxy()
            {
                return base.GetTransparentProxy();
            }

            private string GetCell(string sheetName, int row, int column)
            {
                return cells.TryGetValue($"{sheetName}!{row}!{column}", out var value) ? value : string.Empty;
            }

            private int GetLastUsedColumn(string sheetName)
            {
                var prefix = sheetName + "!";
                return cells.Keys
                    .Where(key => key.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    .Select(key => int.Parse(key.Split('!')[2]))
                    .DefaultIfEmpty(0)
                    .Max();
            }
        }

        private sealed class NoopDisposable : IDisposable
        {
            public void Dispose()
            {
            }
        }
    }
}
```

- [ ] **Step 2: Run the header scanner tests and verify they fail**

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~WorksheetHeaderScannerTests`

Expected: FAIL with missing type `WorksheetHeaderScanner`.

- [ ] **Step 3: Add the header scanner**

Create `src/OfficeAgent.ExcelAddIn/Excel/WorksheetHeaderScanner.cs`:

```csharp
using System;
using System.Collections.Generic;
using OfficeAgent.Core.Models;

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal sealed class WorksheetHeaderScanner
    {
        public AiColumnMappingActualHeader[] Scan(string sheetName, SheetBinding binding, IWorksheetGridAdapter grid)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                throw new ArgumentException("Sheet name is required.", nameof(sheetName));
            }

            if (binding == null)
            {
                throw new ArgumentNullException(nameof(binding));
            }

            if (grid == null)
            {
                throw new ArgumentNullException(nameof(grid));
            }

            var result = new List<AiColumnMappingActualHeader>();
            var headerStartRow = binding.HeaderStartRow <= 0 ? 1 : binding.HeaderStartRow;
            var headerRowCount = binding.HeaderRowCount <= 0 ? 1 : binding.HeaderRowCount;
            var lastUsedColumn = grid.GetLastUsedColumn(sheetName);
            var currentParent = string.Empty;

            for (var column = 1; column <= lastUsedColumn; column++)
            {
                var topText = grid.GetCellText(sheetName, headerStartRow, column) ?? string.Empty;
                var bottomText = headerRowCount > 1
                    ? grid.GetCellText(sheetName, headerStartRow + 1, column) ?? string.Empty
                    : string.Empty;

                if (!string.IsNullOrWhiteSpace(topText))
                {
                    currentParent = topText;
                }

                var actualL1 = headerRowCount > 1 && !string.IsNullOrWhiteSpace(bottomText)
                    ? currentParent
                    : topText;
                var actualL2 = headerRowCount > 1 && !string.IsNullOrWhiteSpace(bottomText)
                    ? bottomText
                    : string.Empty;

                if (string.IsNullOrWhiteSpace(actualL1) && string.IsNullOrWhiteSpace(actualL2))
                {
                    continue;
                }

                result.Add(new AiColumnMappingActualHeader
                {
                    ExcelColumn = column,
                    ActualL1 = actualL1,
                    ActualL2 = actualL2,
                    DisplayText = string.IsNullOrWhiteSpace(actualL2) ? actualL1 : actualL1 + "/" + actualL2,
                });
            }

            return result.ToArray();
        }
    }
}
```

Add this compile item to `src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj` near the other Excel files:

```xml
<Compile Include="Excel\WorksheetHeaderScanner.cs" />
```

- [ ] **Step 4: Run the header scanner tests and verify they pass**

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter FullyQualifiedName~WorksheetHeaderScannerTests`

Expected: PASS.

- [ ] **Step 5: Write the failing execution service tests**

Add tests to `tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs`:

```csharp
[Fact]
public void PrepareAiColumnMappingPreviewScansFullHeaderAreaAndCallsClient()
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
    metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");
    var aiClient = new FakeAiColumnMappingClient
    {
        Response = new AiColumnMappingResponse
        {
            Mappings = new[]
            {
                new AiColumnMappingSuggestion { ExcelColumn = 2, ActualL1 = "项目负责人", TargetHeaderId = "owner_name", TargetApiFieldKey = "owner_name", Confidence = 0.92 },
            },
        },
    };
    var (service, grid) = CreateService(connector, metadataStore, new FakeWorksheetSelectionReader(), aiClient);
    grid.SetCell("Sheet1", 3, 1, "ID");
    grid.SetCell("Sheet1", 3, 2, "项目负责人");
    grid.SetCell("Sheet1", 3, 3, "测试活动111");
    grid.SetCell("Sheet1", 4, 3, "开始时间");

    var preview = (AiColumnMappingPreview)InvokePrivate(service, "PrepareAiColumnMappingPreview", "Sheet1");

    Assert.Equal("项目负责人", aiClient.LastRequest.ActualHeaders.Single(header => header.ExcelColumn == 2).ActualL1);
    Assert.Equal("测试活动111", aiClient.LastRequest.ActualHeaders.Single(header => header.ExcelColumn == 3).ActualL1);
    Assert.Equal("开始时间", aiClient.LastRequest.ActualHeaders.Single(header => header.ExcelColumn == 3).ActualL2);
    Assert.Contains(preview.Items, item => item.TargetHeaderId == "owner_name" && item.Status == AiColumnMappingPreviewStatuses.Accepted);
}

[Fact]
public void ApplyAiColumnMappingPreviewSavesOnlyConfirmedMetadataRows()
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
        HeaderRowCount = 1,
        DataStartRow = 4,
    };
    metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");
    var (service, _) = CreateService(connector, metadataStore, new FakeWorksheetSelectionReader(), new FakeAiColumnMappingClient());
    var preview = new AiColumnMappingPreview
    {
        Items = new[]
        {
            new AiColumnMappingPreviewItem
            {
                ExcelColumn = 2,
                ActualL1 = "项目负责人",
                TargetHeaderId = "owner_name",
                TargetApiFieldKey = "owner_name",
                SuggestedExcelL1 = "项目负责人",
                Confidence = 0.91,
                Status = AiColumnMappingPreviewStatuses.Accepted,
            },
        },
    };

    var result = (AiColumnMappingApplyResult)InvokePrivate(service, "ApplyAiColumnMappingPreview", "Sheet1", preview);

    Assert.Equal(1, result.AppliedCount);
    Assert.Equal("项目负责人", metadataStore.LastSavedFieldMappings.Single(row => row.Values["HeaderId"] == "owner_name").Values["CurrentL1"]);
    Assert.Equal("row_id", metadataStore.LastSavedFieldMappings.Single(row => row.Values["HeaderId"] == "row_id").Values["CurrentL1"]);
}
```

Add helper `FakeAiColumnMappingClient : IAiColumnMappingClient`, overload `CreateService(..., IAiColumnMappingClient aiClient)`, and a generic `InvokePrivate()` helper. Keep the existing overloads delegating to the new one with `null`.

- [ ] **Step 6: Run the execution service tests and verify they fail**

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~PrepareAiColumnMappingPreviewScansFullHeaderAreaAndCallsClient|FullyQualifiedName~ApplyAiColumnMappingPreviewSavesOnlyConfirmedMetadataRows"`

Expected: FAIL with missing methods `PrepareAiColumnMappingPreview` and `ApplyAiColumnMappingPreview`.

- [ ] **Step 7: Wire AI mapping into `WorksheetSyncExecutionService`**

Modify `src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs`:

- Add fields:

```csharp
private readonly IWorksheetMetadataStore metadataStore;
private readonly IAiColumnMappingClient aiColumnMappingClient;
private readonly AiColumnMappingService aiColumnMappingService;
private readonly WorksheetHeaderScanner headerScanner;
```

- Preserve existing constructors by delegating to a new constructor overload that accepts `IAiColumnMappingClient`.
- Store `metadataStore` instead of discarding it.
- Initialize `aiColumnMappingService = new AiColumnMappingService();` and `headerScanner = new WorksheetHeaderScanner();`.
- Add methods:

```csharp
public AiColumnMappingPreview PrepareAiColumnMappingPreview(string sheetName)
{
    if (aiColumnMappingClient == null)
    {
        throw new InvalidOperationException("AI column mapping is not configured.");
    }

    var context = LoadSheetContext(sheetName);
    var actualHeaders = headerScanner.Scan(sheetName, context.Binding, gridAdapter);
    if (actualHeaders.Length == 0)
    {
        throw new InvalidOperationException("No header text was found in the configured header area. Check HeaderStartRow and HeaderRowCount.");
    }

    var request = aiColumnMappingService.BuildRequest(sheetName, context.Definition, context.Mappings, actualHeaders);
    var response = aiColumnMappingClient.Map(request);
    return aiColumnMappingService.CreatePreview(request, response, context.Binding.HeaderRowCount);
}

public AiColumnMappingApplyResult ApplyAiColumnMappingPreview(string sheetName, AiColumnMappingPreview preview)
{
    var context = LoadSheetContext(sheetName);
    var result = aiColumnMappingService.ApplyConfirmedPreview(
        sheetName,
        context.Definition,
        context.Mappings,
        preview,
        context.Binding.HeaderRowCount);
    metadataStore.SaveFieldMappings(sheetName, context.Definition, result.Rows);
    return result;
}
```

Modify `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs` to pass the real client:

```csharp
WorksheetSyncExecutionService = new WorksheetSyncExecutionService(
    WorksheetSyncService,
    WorksheetMetadataStore,
    new ExcelVisibleSelectionReader(Application),
    worksheetGridAdapter,
    new SyncOperationPreviewFactory(),
    WorksheetChangeLogStore,
    WorksheetPendingEditTracker,
    new AiColumnMappingClient(() => SettingsStore.Load()));
```

- [ ] **Step 8: Run the execution service tests and verify they pass**

Run: `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~PrepareAiColumnMappingPreviewScansFullHeaderAreaAndCallsClient|FullyQualifiedName~ApplyAiColumnMappingPreviewSavesOnlyConfirmedMetadataRows"`

Expected: PASS.

- [ ] **Step 9: Commit Task 3**

```powershell
git add -- src/OfficeAgent.ExcelAddIn/Excel/WorksheetHeaderScanner.cs src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj src/OfficeAgent.ExcelAddIn/WorksheetSyncExecutionService.cs src/OfficeAgent.ExcelAddIn/ThisAddIn.cs tests/OfficeAgent.ExcelAddIn.Tests/WorksheetHeaderScannerTests.cs tests/OfficeAgent.ExcelAddIn.Tests/WorksheetSyncExecutionServiceTests.cs
git commit -m "feat: prepare ai column mapping preview"
```

### Task 4: Ribbon Button, Preview Dialog, and Localized Controller Flow

**Files:**
- Create: `src/OfficeAgent.ExcelAddIn/Dialogs/AiColumnMappingPreviewDialog.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/Dialogs/OperationResultDialog.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj`
- Modify: `src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/AgentRibbon.Designer.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/AgentRibbon.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/Localization/HostLocalizedStrings.cs`
- Test: `tests/OfficeAgent.ExcelAddIn.Tests/RibbonSyncControllerTests.cs`
- Test: `tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs`
- Test: `tests/OfficeAgent.ExcelAddIn.Tests/HostLocalizedStringsTests.cs`

- [ ] **Step 1: Write failing Ribbon and localization tests**

Add to `tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs`:

```csharp
[Fact]
public void AiMapColumnsButtonIsPlacedUnderInitializeSheetInProjectGroup()
{
    var designerText = File.ReadAllText(ResolveRepositoryPath("src", "OfficeAgent.ExcelAddIn", "AgentRibbon.Designer.cs"));

    var initializeIndex = designerText.IndexOf("this.groupProject.Items.Add(this.initializeSheetButton);", StringComparison.Ordinal);
    var aiMapIndex = designerText.IndexOf("this.groupProject.Items.Add(this.aiMapColumnsButton);", StringComparison.Ordinal);

    Assert.True(initializeIndex >= 0);
    Assert.True(aiMapIndex > initializeIndex);
    Assert.Contains("this.aiMapColumnsButton.Label = \"AI map columns\";", designerText, StringComparison.Ordinal);
    Assert.Contains("this.aiMapColumnsButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular;", designerText, StringComparison.Ordinal);
}
```

Add to `tests/OfficeAgent.ExcelAddIn.Tests/HostLocalizedStringsTests.cs`:

```csharp
[Theory]
[InlineData("zh", "AI映射列")]
[InlineData("en", "AI map columns")]
public void ForLocaleReturnsAiColumnMappingRibbonLabel(string locale, string expectedLabel)
{
    var strings = CreateStrings(locale);

    Assert.Equal(expectedLabel, GetString(strings, "RibbonAiMapColumnsButtonLabel"));
}
```

Add to `tests/OfficeAgent.ExcelAddIn.Tests/RibbonSyncControllerTests.cs`:

```csharp
[Fact]
public void ExecuteAiColumnMappingConfirmsPreviewBeforeSavingMappings()
{
    var connector = new FakeSystemConnector();
    var metadataStore = new FakeWorksheetMetadataStore();
    var dialogService = new FakeDialogService { AiColumnMappingConfirmResult = true };
    metadataStore.Bindings["Sheet1"] = new SheetBinding
    {
        SheetName = "Sheet1",
        SystemKey = "current-business-system",
        ProjectId = "performance",
        ProjectName = "绩效项目",
        HeaderStartRow = 3,
        HeaderRowCount = 1,
        DataStartRow = 4,
    };
    metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");
    var aiClient = new FakeAiColumnMappingClient
    {
        Response = new AiColumnMappingResponse
        {
            Mappings = new[]
            {
                new AiColumnMappingSuggestion { ExcelColumn = 2, ActualL1 = "项目负责人", TargetHeaderId = "owner_name", TargetApiFieldKey = "owner_name", Confidence = 0.91 },
            },
        },
    };
    var controller = CreateController(connector, metadataStore, dialogService, () => "Sheet1", authenticationLoginAction: null, aiClient: aiClient);
    InvokeRefresh(controller);

    Invoke(controller, "ExecuteAiColumnMapping");

    Assert.Single(dialogService.AiColumnMappingPreviews);
    Assert.Equal("项目负责人", metadataStore.LastSavedFieldMappings.Single(row => row.Values["HeaderId"] == "owner_name").Values["CurrentL1"]);
    Assert.Contains(dialogService.InfoMessages, message => message.Contains("1", StringComparison.Ordinal));
}
```

Update `FakeDialogService` to track `AiColumnMappingPreviews`, return `AiColumnMappingConfirmResult` for `ConfirmAiColumnMapping`, and update the controller helper to accept `IAiColumnMappingClient`.

- [ ] **Step 2: Run the tests and verify they fail**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~AiMapColumnsButtonIsPlacedUnderInitializeSheetInProjectGroup|FullyQualifiedName~ForLocaleReturnsAiColumnMappingRibbonLabel|FullyQualifiedName~ExecuteAiColumnMappingConfirmsPreviewBeforeSavingMappings"
```

Expected: FAIL with missing label/button/controller/dialog members.

- [ ] **Step 3: Add localized strings**

Modify `src/OfficeAgent.ExcelAddIn/Localization/HostLocalizedStrings.cs`:

```csharp
public string RibbonAiMapColumnsButtonLabel => Locale == "zh" ? "AI映射列" : "AI map columns";

public string AiColumnMappingPreviewDialogTitle => Locale == "zh" ? "AI映射列预览" : "AI column mapping preview";

public string AiColumnMappingCompletedMessage(int appliedCount, int skippedCount)
{
    return Locale == "zh"
        ? $"AI映射列完成。已写入：{appliedCount}，未写入：{skippedCount}。"
        : $"AI column mapping completed. Applied: {appliedCount}; skipped: {skippedCount}.";
}

public string AiColumnMappingNoAcceptedMappingsMessage => Locale == "zh"
    ? "没有可写入的高置信度映射。"
    : "No high-confidence mappings are available to write.";

public string AiColumnMappingPreviewInstructionText => Locale == "zh"
    ? "确认后只会更新 xISDP_Setting 中的 Excel L1 / Excel L2，不会修改业务单元格。"
    : "Confirming updates only Excel L1 / Excel L2 in xISDP_Setting. Business cells are not changed.";
```

- [ ] **Step 4: Extend the dialog service and add the preview dialog**

Modify `src/OfficeAgent.ExcelAddIn/Dialogs/OperationResultDialog.cs`:

```csharp
bool ConfirmAiColumnMapping(AiColumnMappingPreview preview);
```

Implement it in `RibbonSyncDialogService`:

```csharp
public bool ConfirmAiColumnMapping(AiColumnMappingPreview preview)
{
    return AiColumnMappingPreviewDialog.Confirm(preview);
}
```

Create `src/OfficeAgent.ExcelAddIn/Dialogs/AiColumnMappingPreviewDialog.cs`:

```csharp
using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using OfficeAgent.Core.Models;
using OfficeAgent.ExcelAddIn.Localization;

namespace OfficeAgent.ExcelAddIn.Dialogs
{
    internal sealed class AiColumnMappingPreviewDialog : Form
    {
        public static bool Confirm(AiColumnMappingPreview preview)
        {
            using (var dialog = new AiColumnMappingPreviewDialog(preview, Globals.ThisAddIn?.HostLocalizedStrings ?? HostLocalizedStrings.ForLocale("en")))
            {
                return dialog.ShowDialog() == DialogResult.OK;
            }
        }

        public AiColumnMappingPreviewDialog(AiColumnMappingPreview preview, HostLocalizedStrings strings)
        {
            var localizedStrings = strings ?? HostLocalizedStrings.ForLocale("en");
            Text = localizedStrings.AiColumnMappingPreviewDialogTitle;
            StartPosition = FormStartPosition.CenterParent;
            Size = new Size(860, 520);
            MinimumSize = new Size(720, 420);
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = true;
            MinimizeBox = false;
            ShowInTaskbar = false;

            var instruction = new Label
            {
                Dock = DockStyle.Top,
                Height = 44,
                Padding = new Padding(12, 10, 12, 6),
                Text = localizedStrings.AiColumnMappingPreviewInstructionText,
            };

            var grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            };
            grid.Columns.Add("ExcelColumn", "Column");
            grid.Columns.Add("ActualHeader", "Actual header");
            grid.Columns.Add("TargetHeader", "ISDP header");
            grid.Columns.Add("SuggestedExcelHeader", "Excel L1 / L2");
            grid.Columns.Add("Confidence", "Confidence");
            grid.Columns.Add("Status", "Status");
            grid.Columns.Add("Reason", "Reason");

            foreach (var item in preview?.Items ?? Array.Empty<AiColumnMappingPreviewItem>())
            {
                grid.Rows.Add(
                    item.ExcelColumn,
                    FormatHeader(item.ActualL1, item.ActualL2),
                    FormatHeader(item.TargetIsdpL1, item.TargetIsdpL2),
                    FormatHeader(item.SuggestedExcelL1, item.SuggestedExcelL2),
                    item.Confidence.ToString("0.00"),
                    item.Status,
                    item.Reason);
            }

            var okButton = new Button { Text = localizedStrings.OkButtonText, DialogResult = DialogResult.OK, AutoSize = true };
            var cancelButton = new Button { Text = localizedStrings.CancelButtonText, DialogResult = DialogResult.Cancel, AutoSize = true };
            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                FlowDirection = FlowDirection.RightToLeft,
                Height = 48,
                Padding = new Padding(12, 8, 12, 8),
            };
            buttons.Controls.Add(cancelButton);
            buttons.Controls.Add(okButton);

            AcceptButton = okButton;
            CancelButton = cancelButton;
            Controls.Add(grid);
            Controls.Add(instruction);
            Controls.Add(buttons);
        }

        private static string FormatHeader(string l1, string l2)
        {
            return string.IsNullOrWhiteSpace(l2) ? l1 ?? string.Empty : (l1 ?? string.Empty) + "/" + l2;
        }
    }
}
```

Add to `src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj`:

```xml
<Compile Include="Dialogs\AiColumnMappingPreviewDialog.cs" />
```

- [ ] **Step 5: Add the controller flow**

Modify `src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs`:

```csharp
public void ExecuteAiColumnMapping()
{
    if (!EnsureProjectSelected())
    {
        return;
    }

    try
    {
        var strings = GetStrings();
        var sheetName = GetRequiredSheetName();
        var preview = EnsureExecutionService().PrepareAiColumnMappingPreview(sheetName);
        if (!dialogService.ConfirmAiColumnMapping(preview))
        {
            return;
        }

        var result = executionService.ApplyAiColumnMappingPreview(sheetName, preview);
        dialogService.ShowInfo(result.AppliedCount == 0
            ? strings.AiColumnMappingNoAcceptedMappingsMessage
            : strings.AiColumnMappingCompletedMessage(result.AppliedCount, result.SkippedCount));
    }
    catch (Exception ex)
    {
        dialogService.ShowError(ex.Message);
    }
}
```

- [ ] **Step 6: Add the Ribbon button**

Modify `src/OfficeAgent.ExcelAddIn/AgentRibbon.Designer.cs`:

- Create `aiMapColumnsButton` after `initializeSheetButton`.
- Add it to `groupProject` immediately after `initializeSheetButton`.
- Set:

```csharp
this.aiMapColumnsButton.Label = "AI map columns";
this.aiMapColumnsButton.Name = "aiMapColumnsButton";
this.aiMapColumnsButton.OfficeImageId = "TableAutoFormat";
this.aiMapColumnsButton.ShowImage = true;
this.aiMapColumnsButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular;
this.aiMapColumnsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AiMapColumnsButton_Click);
```

Modify `src/OfficeAgent.ExcelAddIn/AgentRibbon.cs`:

```csharp
private void AiMapColumnsButton_Click(object sender, RibbonControlEventArgs e)
{
    Globals.ThisAddIn.RibbonSyncController?.ExecuteAiColumnMapping();
}
```

In `ApplyLocalizedLabels()`:

```csharp
aiMapColumnsButton.Label = FormatRibbonButtonLabel(strings.RibbonAiMapColumnsButtonLabel);
```

- [ ] **Step 7: Run the Ribbon/localization/controller tests and verify they pass**

Run:

```powershell
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj --filter "FullyQualifiedName~AiMapColumnsButtonIsPlacedUnderInitializeSheetInProjectGroup|FullyQualifiedName~ForLocaleReturnsAiColumnMappingRibbonLabel|FullyQualifiedName~ExecuteAiColumnMappingConfirmsPreviewBeforeSavingMappings"
```

Expected: PASS.

- [ ] **Step 8: Commit Task 4**

```powershell
git add -- src/OfficeAgent.ExcelAddIn/Dialogs/AiColumnMappingPreviewDialog.cs src/OfficeAgent.ExcelAddIn/Dialogs/OperationResultDialog.cs src/OfficeAgent.ExcelAddIn/OfficeAgent.ExcelAddIn.csproj src/OfficeAgent.ExcelAddIn/RibbonSyncController.cs src/OfficeAgent.ExcelAddIn/AgentRibbon.Designer.cs src/OfficeAgent.ExcelAddIn/AgentRibbon.cs src/OfficeAgent.ExcelAddIn/Localization/HostLocalizedStrings.cs tests/OfficeAgent.ExcelAddIn.Tests/RibbonSyncControllerTests.cs tests/OfficeAgent.ExcelAddIn.Tests/AgentRibbonConfigurationTests.cs tests/OfficeAgent.ExcelAddIn.Tests/HostLocalizedStringsTests.cs
git commit -m "feat: add ai column mapping ribbon action"
```

### Task 5: Documentation and Full Verification

**Files:**
- Modify: `docs/modules/ribbon-sync-current-behavior.md`
- Modify: `docs/ribbon-sync-real-system-integration-guide.md`
- Modify: `docs/vsto-manual-test-checklist.md`

- [ ] **Step 1: Update Ribbon Sync current behavior**

In `docs/modules/ribbon-sync-current-behavior.md`, update:

- Section 2 Ribbon entry list: add `AI映射列 / AI map columns` under `项目 / Project`, below `初始化当前表 / Initialize sheet`.
- Section 3.3 `SheetFieldMappings`: add that AI mapping updates only `Excel L1 / Excel L2`.
- Section 5 table header matching: add that users can now run AI mapping to write current header names into metadata after preview confirmation.
- Section 9 code entry list: add `AiColumnMappingClient`, `AiColumnMappingService`, and `WorksheetHeaderScanner`.
- Section 10 test entry list: add the new test files.

- [ ] **Step 2: Update the real-system integration guide**

In `docs/ribbon-sync-real-system-integration-guide.md`, update Section 7.1 / 7.2:

```markdown
如果用户只改了 Excel 可见表头，当前仍不会在上传 / 下载时静默猜测。但用户可以点击 `AI映射列`，让插件扫描当前绑定 sheet 的完整表头区，并在确认预览后把匹配成功项写入 `SheetFieldMappings.Excel L1 / Excel L2`。

该能力不改 `ISDP L1 / ISDP L2`、接口字段身份列或业务单元格。真实系统连接器仍需提供稳定的 `HeaderId / ApiFieldKey` 和清晰的默认表头文本，供 AI 映射候选列表使用。
```

- [ ] **Step 3: Update the manual test checklist**

In `docs/vsto-manual-test-checklist.md`, add a Ribbon Sync manual case:

```markdown
### AI 映射列

1. 选择项目并执行 `初始化当前表`。
2. 在业务 sheet 中把表头改成与 `ISDP L1 / ISDP L2` 不同但语义相近的名称，例如把 `负责人` 改成 `项目负责人`。
3. 点击 `AI映射列`。
4. 确认预览中显示实际表头、建议映射字段、置信度和未匹配项。
5. 点击确认后，检查 `xISDP_Setting.SheetFieldMappings` 只更新 `Excel L1 / Excel L2`。
6. 执行部分上传或下载，确认新表头可以被识别。
7. 重复一次并在预览中点击取消，确认 `xISDP_Setting` 不变化。
```

- [ ] **Step 4: Run full targeted verification**

Run:

```powershell
dotnet test tests/OfficeAgent.Core.Tests/OfficeAgent.Core.Tests.csproj
dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj
dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj
```

Expected: all three commands PASS. If the VSTO test project cannot locate Visual Studio MSBuild, report that as an environment blocker and run the Core and Infrastructure tests at minimum.

- [ ] **Step 5: Run whitespace and status checks**

Run:

```powershell
git diff --check
git status --short
```

Expected: `git diff --check` exits 0. `git status --short` shows only intended feature files plus the pre-existing unstaged `src/OfficeAgent.ExcelAddIn/Properties/Version.g.cs` if it is still present.

- [ ] **Step 6: Commit Task 5**

```powershell
git add -- docs/modules/ribbon-sync-current-behavior.md docs/ribbon-sync-real-system-integration-guide.md docs/vsto-manual-test-checklist.md
git commit -m "docs: document ai column mapping"
```

## Final Verification Before Completion

- [ ] Run `dotnet test tests/OfficeAgent.Core.Tests/OfficeAgent.Core.Tests.csproj`.
- [ ] Run `dotnet test tests/OfficeAgent.Infrastructure.Tests/OfficeAgent.Infrastructure.Tests.csproj`.
- [ ] Run `dotnet test tests/OfficeAgent.ExcelAddIn.Tests/OfficeAgent.ExcelAddIn.Tests.csproj`.
- [ ] Run `git diff --check`.
- [ ] Run `git status --short --branch`.
- [ ] Confirm `src/OfficeAgent.ExcelAddIn/Properties/Version.g.cs` was not staged or reverted unless the user explicitly requested it.

## Spec Coverage Self-Review

- Ribbon independent button under `Initialize sheet`: Task 4.
- Full header-area scan rather than selection scan: Task 3.
- Reuse existing `BASE_URL / API_KEY / MODEL`: Task 2 and Task 3 injection.
- Preview confirmation before write: Task 4.
- Write only `Excel L1 / Excel L2`, preserve identity and default fields: Task 1 and Task 3.
- Unmatched or low-confidence items remain unchanged: Task 1.
- Native WinForms, no task pane changes: Task 4.
- Documentation updates: Task 5.
