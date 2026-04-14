using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
using OfficeAgent.Core.Sync;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class WorksheetSyncExecutionServiceTests
    {
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
            Assert.Equal("performance", connector.LastFieldMappingDefinitionProjectId);
            Assert.NotEmpty(metadataStore.LastSavedFieldMappings);
        }

        [Fact]
        public void ExecuteFullDownloadHonorsConfiguredHeaderAndDataRowsWhenSheetHeadersAreEmpty()
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
            Assert.Equal("项目负责人", grid.GetCell("Sheet1", 3, 2));
            Assert.Equal("测试活动111", grid.GetCell("Sheet1", 3, 3));
            Assert.Equal("开始时间", grid.GetCell("Sheet1", 4, 3));
            Assert.Equal("结束时间", grid.GetCell("Sheet1", 4, 4));
            Assert.Equal("row-1", grid.GetCell("Sheet1", 6, 1));
            Assert.Equal("张三", grid.GetCell("Sheet1", 6, 2));
            Assert.Equal("2026-01-02", grid.GetCell("Sheet1", 6, 3));
            Assert.Equal("2026-01-05", grid.GetCell("Sheet1", 6, 4));

            Assert.Contains(grid.Merges, merge => merge.SheetName == "Sheet1" && merge.Row == 3 && merge.Column == 1 && merge.RowSpan == 2 && merge.ColumnSpan == 1);
            Assert.Contains(grid.Merges, merge => merge.SheetName == "Sheet1" && merge.Row == 3 && merge.Column == 3 && merge.RowSpan == 1 && merge.ColumnSpan == 2);
        }

        [Fact]
        public void ExecutePartialDownloadUsesRecognizedHeadersAndIdLookupOutsideSelection()
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
            connector.FindResult = new[] { CreateRow("row-1", "张三", "2026-02-01", "2026-02-09") };

            var selectionReader = new FakeWorksheetSelectionReader
            {
                VisibleCells = new[]
                {
                    new SelectedVisibleCell { Row = 6, Column = 3, Value = "旧开始时间" },
                },
            };
            var (service, grid) = CreateService(connector, metadataStore, selectionReader);
            SeedRecognizedHeaders(grid, "Sheet1", binding);
            grid.SetCell("Sheet1", 6, 1, "row-1");
            grid.SetCell("Sheet1", 6, 2, "张三");
            grid.SetCell("Sheet1", 6, 3, "旧开始时间");
            grid.SetCell("Sheet1", 6, 4, "旧结束时间");

            var plan = InvokePrepare(service, "PreparePartialDownload", "Sheet1");
            InvokeExecute(service, "ExecuteDownload", plan);

            Assert.Equal("performance", connector.LastFindProjectId);
            Assert.Equal(new[] { "row-1" }, connector.LastFindRowIds);
            Assert.Equal(new[] { "start_12345678" }, connector.LastFindFieldKeys);
            Assert.Equal("2026-02-01", grid.GetCell("Sheet1", 6, 3));
            Assert.Equal("旧结束时间", grid.GetCell("Sheet1", 6, 4));
        }

        [Fact]
        public void ExecuteFullDownloadDoesNotRewriteExistingRecognizedHeaders()
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
            connector.FindResult = new[] { CreateRow("row-1", "张三", "2026-01-02", "2026-01-05") };

            var (service, grid) = CreateService(connector, metadataStore, new FakeWorksheetSelectionReader());
            SeedRecognizedHeaders(grid, "Sheet1", binding);
            grid.SetCell("Sheet1", 6, 1, "row-1");
            grid.SetCell("Sheet1", 6, 2, "旧负责人");
            grid.SetCell("Sheet1", 6, 3, "旧开始");
            grid.SetCell("Sheet1", 6, 4, "旧结束");

            var plan = InvokePrepare(service, "PrepareFullDownload", "Sheet1");
            InvokeExecute(service, "ExecuteDownload", plan);

            Assert.DoesNotContain(grid.ClearedRanges, range => range.StartRow <= 4 && range.EndRow >= 3);
            Assert.Empty(grid.Merges);
            Assert.Equal("ID", grid.GetCell("Sheet1", 3, 1));
            Assert.Equal("项目负责人", grid.GetCell("Sheet1", 3, 2));
            Assert.Equal("测试活动111", grid.GetCell("Sheet1", 3, 3));
            Assert.Equal("开始时间", grid.GetCell("Sheet1", 4, 3));
            Assert.Equal("2026-01-02", grid.GetCell("Sheet1", 6, 3));
        }

        [Fact]
        public void ExecuteFullUploadUsesConfiguredDataStartRowAndRecognizedColumns()
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

            var (service, grid) = CreateService(connector, metadataStore, new FakeWorksheetSelectionReader());
            SeedRecognizedHeaders(grid, "Sheet1", binding);
            grid.SetCell("Sheet1", 5, 1, "统计行");
            grid.SetCell("Sheet1", 6, 1, "row-1");
            grid.SetCell("Sheet1", 6, 2, "李四");
            grid.SetCell("Sheet1", 6, 3, "2026-01-02");
            grid.SetCell("Sheet1", 6, 4, "2026-01-05");
            grid.SetCell("Sheet1", 7, 1, string.Empty);
            grid.SetCell("Sheet1", 7, 2, "无ID");
            grid.SetCell("Sheet1", 7, 3, "2026-03-01");
            grid.SetCell("Sheet1", 7, 4, "2026-03-05");

            var plan = InvokePrepare(service, "PrepareFullUpload", "Sheet1");
            var preview = ReadPreview(plan);
            Assert.Equal(3, preview.Changes.Length);

            InvokeExecute(service, "ExecuteUpload", plan);

            Assert.Equal("performance", connector.LastBatchSaveProjectId);
            Assert.Equal(3, connector.LastBatchSaveChanges.Count);
            Assert.Contains(connector.LastBatchSaveChanges, change => change.RowId == "row-1" && change.ApiFieldKey == "owner_name" && change.NewValue == "李四");
            Assert.Contains(connector.LastBatchSaveChanges, change => change.RowId == "row-1" && change.ApiFieldKey == "start_12345678" && change.NewValue == "2026-01-02");
            Assert.Contains(connector.LastBatchSaveChanges, change => change.RowId == "row-1" && change.ApiFieldKey == "end_12345678" && change.NewValue == "2026-01-05");
            Assert.DoesNotContain(connector.LastBatchSaveChanges, change => string.IsNullOrWhiteSpace(change.RowId));
        }

        [Fact]
        public void ExecuteFullDownloadAutoReinitializesWhenStoredMappingsLackUsableIdDefinition()
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
            metadataStore.FieldMappings["Sheet1"] = BuildLegacyMappingsWithoutIdFlag("Sheet1");
            connector.FindResult = new[] { CreateRow("row-1", "张三", "2026-01-02", "2026-01-05") };

            var (service, grid) = CreateService(connector, metadataStore, new FakeWorksheetSelectionReader());

            var plan = InvokePrepare(service, "PrepareFullDownload", "Sheet1");
            InvokeExecute(service, "ExecuteDownload", plan);

            Assert.Equal(3, metadataStore.LastSavedBinding.HeaderStartRow);
            Assert.Equal(2, metadataStore.LastSavedBinding.HeaderRowCount);
            Assert.Equal(6, metadataStore.LastSavedBinding.DataStartRow);
            Assert.Contains(
                metadataStore.LastSavedFieldMappings,
                row => string.Equals(row.Values["ApiFieldKey"], "row_id", StringComparison.Ordinal) &&
                       string.Equals(row.Values["IsIdColumn"], "true", StringComparison.OrdinalIgnoreCase));
            Assert.Equal("row-1", grid.GetCell("Sheet1", 6, 1));
            Assert.Equal("张三", grid.GetCell("Sheet1", 6, 2));
        }

        [Fact]
        public void ExecutePartialUploadUsesRecognizedHeadersAndIdLookupOutsideSelection()
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

            var selectionReader = new FakeWorksheetSelectionReader
            {
                VisibleCells = new[]
                {
                    new SelectedVisibleCell { Row = 6, Column = 4, Value = "2026-01-10" },
                },
            };
            var (service, grid) = CreateService(connector, metadataStore, selectionReader);
            SeedRecognizedHeaders(grid, "Sheet1", binding);
            grid.SetCell("Sheet1", 6, 1, "row-1");
            grid.SetCell("Sheet1", 6, 2, "张三");
            grid.SetCell("Sheet1", 6, 3, "2026-01-02");
            grid.SetCell("Sheet1", 6, 4, "2026-01-10");

            var plan = InvokePrepare(service, "PreparePartialUpload", "Sheet1");
            var preview = ReadPreview(plan);
            Assert.Single(preview.Changes);
            Assert.Equal("row-1", preview.Changes[0].RowId);
            Assert.Equal("end_12345678", preview.Changes[0].ApiFieldKey);

            InvokeExecute(service, "ExecuteUpload", plan);

            Assert.Equal("performance", connector.LastBatchSaveProjectId);
            Assert.Single(connector.LastBatchSaveChanges);
            Assert.Equal("2026-01-10", connector.LastBatchSaveChanges[0].NewValue);
        }

        private static (object Service, FakeWorksheetGridAdapter Grid) CreateService(
            FakeSystemConnector connector,
            FakeWorksheetMetadataStore metadataStore,
            FakeWorksheetSelectionReader selectionReader)
        {
            var assembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var serviceType = assembly.GetType("OfficeAgent.ExcelAddIn.WorksheetSyncExecutionService", throwOnError: true);
            var gridInterface = assembly.GetType("OfficeAgent.ExcelAddIn.Excel.IWorksheetGridAdapter", throwOnError: true);
            var grid = new FakeWorksheetGridAdapter(gridInterface);
            var syncService = new WorksheetSyncService(
                connector,
                metadataStore,
                new WorksheetChangeTracker(),
                new SyncOperationPreviewFactory());

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

            var service = ctor.Invoke(new object[]
            {
                syncService,
                metadataStore,
                selectionReader,
                grid.GetTransparentProxy(),
                new SyncOperationPreviewFactory(),
            });

            return (service, grid);
        }

        private static void SeedRecognizedHeaders(FakeWorksheetGridAdapter grid, string sheetName, SheetBinding binding)
        {
            var row = binding.HeaderStartRow;
            grid.SetCell(sheetName, row, 1, "ID");
            grid.SetCell(sheetName, row, 2, "项目负责人");
            grid.SetCell(sheetName, row, 3, "测试活动111");

            if (binding.HeaderRowCount > 1)
            {
                grid.SetCell(sheetName, row + 1, 3, "开始时间");
                grid.SetCell(sheetName, row + 1, 4, "结束时间");
            }
        }

        private static object InvokePrepare(object service, string methodName, string sheetName)
        {
            var method = service.GetType().GetMethod(
                methodName,
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

            if (method == null)
            {
                throw new InvalidOperationException($"{methodName} was not found.");
            }

            return method.Invoke(service, new object[] { sheetName });
        }

        private static void InvokeInitialize(object service, string sheetName, ProjectOption project)
        {
            var method = service.GetType().GetMethod(
                "InitializeCurrentSheet",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

            if (method == null)
            {
                throw new InvalidOperationException("InitializeCurrentSheet was not found.");
            }

            method.Invoke(service, new object[] { sheetName, project });
        }

        private static void InvokeExecute(object service, string methodName, object plan)
        {
            var method = service.GetType().GetMethod(
                methodName,
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

            if (method == null)
            {
                throw new InvalidOperationException($"{methodName} was not found.");
            }

            method.Invoke(service, new[] { plan });
        }

        private static SyncOperationPreview ReadPreview(object plan)
        {
            var property = plan.GetType().GetProperty(
                "Preview",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

            if (property == null)
            {
                throw new InvalidOperationException("Preview property was not found.");
            }

            return (SyncOperationPreview)property.GetValue(plan);
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

        private static FieldMappingTableDefinition BuildDefinition()
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

        private static SheetFieldMappingRow[] BuildDefaultMappings(string sheetName)
        {
            return new[]
            {
                CreateMappingRow(sheetName, "row_id", "single", true, currentSingle: "ID"),
                CreateMappingRow(sheetName, "owner_name", "single", false, defaultSingle: "负责人", currentSingle: "项目负责人"),
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

        private static SheetFieldMappingRow[] BuildLegacyMappingsWithoutIdFlag(string sheetName)
        {
            return new[]
            {
                new SheetFieldMappingRow
                {
                    SheetName = sheetName,
                    Values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["HeaderId"] = "row_id",
                        ["CurrentSingleDisplayName"] = "ID",
                    },
                },
                new SheetFieldMappingRow
                {
                    SheetName = sheetName,
                    Values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["HeaderId"] = "owner_name",
                        ["CurrentSingleDisplayName"] = "项目负责人",
                    },
                },
            };
        }

        private static SheetFieldMappingRow CreateMappingRow(
            string sheetName,
            string apiFieldKey,
            string headerType,
            bool isIdColumn,
            string defaultSingle = "",
            string currentSingle = "",
            string defaultParent = "",
            string currentParent = "",
            string defaultChild = "",
            string currentChild = "",
            string activityId = "",
            string propertyId = "")
        {
            return new SheetFieldMappingRow
            {
                SheetName = sheetName,
                Values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["HeaderId"] = apiFieldKey,
                    ["HeaderType"] = headerType,
                    ["ApiFieldKey"] = apiFieldKey,
                    ["IsIdColumn"] = isIdColumn ? "true" : "false",
                    ["DefaultSingleDisplayName"] = defaultSingle,
                    ["CurrentSingleDisplayName"] = currentSingle,
                    ["DefaultParentDisplayName"] = defaultParent,
                    ["CurrentParentDisplayName"] = currentParent,
                    ["DefaultChildDisplayName"] = defaultChild,
                    ["CurrentChildDisplayName"] = currentChild,
                    ["ActivityId"] = activityId,
                    ["PropertyId"] = propertyId,
                },
            };
        }

        private static IDictionary<string, object> CreateRow(string rowId, string ownerName, string start, string end)
        {
            return new Dictionary<string, object>(StringComparer.Ordinal)
            {
                ["row_id"] = rowId,
                ["owner_name"] = ownerName,
                ["start_12345678"] = start,
                ["end_12345678"] = end,
            };
        }

        private sealed class FakeSystemConnector : ISystemConnector
        {
            public FakeSystemConnector()
            {
                BindingSeed = new SheetBinding
                {
                    SheetName = "Sheet1",
                    SystemKey = "current-business-system",
                    ProjectId = "performance",
                    ProjectName = "绩效项目",
                    HeaderStartRow = 1,
                    HeaderRowCount = 2,
                    DataStartRow = 3,
                };
                FieldMappingDefinition = BuildDefinition();
                FieldMappingSeedRows = BuildDefaultMappings("Sheet1");
            }

            public SheetBinding BindingSeed { get; set; }

            public FieldMappingTableDefinition FieldMappingDefinition { get; set; }

            public IReadOnlyList<SheetFieldMappingRow> FieldMappingSeedRows { get; set; }

            public IReadOnlyList<IDictionary<string, object>> FindResult { get; set; } = Array.Empty<IDictionary<string, object>>();

            public string LastFieldMappingDefinitionProjectId { get; private set; }

            public string LastFindProjectId { get; private set; }

            public IReadOnlyList<string> LastFindRowIds { get; private set; } = Array.Empty<string>();

            public IReadOnlyList<string> LastFindFieldKeys { get; private set; } = Array.Empty<string>();

            public string LastBatchSaveProjectId { get; private set; }

            public IReadOnlyList<CellChange> LastBatchSaveChanges { get; private set; } = Array.Empty<CellChange>();

            public IReadOnlyList<ProjectOption> GetProjects()
            {
                return Array.Empty<ProjectOption>();
            }

            public SheetBinding CreateBindingSeed(string sheetName, ProjectOption project)
            {
                return new SheetBinding
                {
                    SheetName = sheetName,
                    SystemKey = project?.SystemKey ?? string.Empty,
                    ProjectId = project?.ProjectId ?? string.Empty,
                    ProjectName = project?.DisplayName ?? string.Empty,
                    HeaderStartRow = BindingSeed.HeaderStartRow,
                    HeaderRowCount = BindingSeed.HeaderRowCount,
                    DataStartRow = BindingSeed.DataStartRow,
                };
            }

            public FieldMappingTableDefinition GetFieldMappingDefinition(string projectId)
            {
                LastFieldMappingDefinitionProjectId = projectId;
                return FieldMappingDefinition;
            }

            public IReadOnlyList<SheetFieldMappingRow> BuildFieldMappingSeed(string sheetName, string projectId)
            {
                return FieldMappingSeedRows;
            }

            public WorksheetSchema GetSchema(string projectId)
            {
                throw new NotSupportedException();
            }

            public IReadOnlyList<IDictionary<string, object>> Find(
                string projectId,
                IReadOnlyList<string> rowIds,
                IReadOnlyList<string> fieldKeys)
            {
                LastFindProjectId = projectId;
                LastFindRowIds = rowIds?.ToArray() ?? Array.Empty<string>();
                LastFindFieldKeys = fieldKeys?.ToArray() ?? Array.Empty<string>();

                IEnumerable<IDictionary<string, object>> rows = FindResult;

                if (LastFindRowIds.Count > 0)
                {
                    rows = rows.Where(row => LastFindRowIds.Contains(Convert.ToString(row["row_id"])));
                }

                if (LastFindFieldKeys.Count > 0)
                {
                    rows = rows.Select(row =>
                    {
                        var projected = new Dictionary<string, object>(StringComparer.Ordinal)
                        {
                            ["row_id"] = row["row_id"],
                        };

                        foreach (var fieldKey in LastFindFieldKeys)
                        {
                            if (row.TryGetValue(fieldKey, out var value))
                            {
                                projected[fieldKey] = value;
                            }
                        }

                        return (IDictionary<string, object>)projected;
                    });
                }

                return rows.ToArray();
            }

            public void BatchSave(string projectId, IReadOnlyList<CellChange> changes)
            {
                LastBatchSaveProjectId = projectId;
                LastBatchSaveChanges = changes?.ToArray() ?? Array.Empty<CellChange>();
            }
        }

        private sealed class FakeWorksheetMetadataStore : IWorksheetMetadataStore
        {
            public Dictionary<string, SheetBinding> Bindings { get; } = new Dictionary<string, SheetBinding>(StringComparer.OrdinalIgnoreCase);

            public Dictionary<string, SheetFieldMappingRow[]> FieldMappings { get; } = new Dictionary<string, SheetFieldMappingRow[]>(StringComparer.OrdinalIgnoreCase);

            public SheetBinding LastSavedBinding { get; private set; }

            public FieldMappingTableDefinition LastSavedFieldMappingDefinition { get; private set; }

            public SheetFieldMappingRow[] LastSavedFieldMappings { get; private set; } = Array.Empty<SheetFieldMappingRow>();

            public void SaveBinding(SheetBinding binding)
            {
                LastSavedBinding = binding;
                Bindings[binding.SheetName] = binding;
            }

            public SheetBinding LoadBinding(string sheetName)
            {
                if (!Bindings.TryGetValue(sheetName, out var binding))
                {
                    throw new InvalidOperationException("No binding.");
                }

                return binding;
            }

            public void SaveFieldMappings(string sheetName, FieldMappingTableDefinition definition, IReadOnlyList<SheetFieldMappingRow> rows)
            {
                LastSavedFieldMappingDefinition = definition;
                LastSavedFieldMappings = (rows ?? Array.Empty<SheetFieldMappingRow>()).ToArray();
                FieldMappings[sheetName] = LastSavedFieldMappings;
            }

            public SheetFieldMappingRow[] LoadFieldMappings(string sheetName, FieldMappingTableDefinition definition)
            {
                return FieldMappings.TryGetValue(sheetName, out var rows)
                    ? rows
                    : Array.Empty<SheetFieldMappingRow>();
            }

            public void ClearFieldMappings(string sheetName)
            {
                FieldMappings.Remove(sheetName);
            }

            public WorksheetSnapshotCell[] LoadSnapshot(string sheetName)
            {
                return Array.Empty<WorksheetSnapshotCell>();
            }

            public void SaveSnapshot(string sheetName, WorksheetSnapshotCell[] cells)
            {
            }
        }

        private sealed class FakeWorksheetSelectionReader : IWorksheetSelectionReader
        {
            public IReadOnlyList<SelectedVisibleCell> VisibleCells { get; set; } = Array.Empty<SelectedVisibleCell>();

            public IReadOnlyList<SelectedVisibleCell> ReadVisibleSelection()
            {
                return VisibleCells;
            }
        }

        private sealed class FakeWorksheetGridAdapter : RealProxy
        {
            private readonly Dictionary<string, string> cells = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            public FakeWorksheetGridAdapter(Type interfaceType)
                : base(interfaceType)
            {
            }

            public List<MergeRecord> Merges { get; } = new List<MergeRecord>();

            public List<ClearRangeRecord> ClearedRanges { get; } = new List<ClearRangeRecord>();

            public override IMessage Invoke(IMessage msg)
            {
                var call = (IMethodCallMessage)msg;

                switch (call.MethodName)
                {
                    case "GetCellText":
                        return new ReturnMessage(GetCell(
                            (string)call.InArgs[0],
                            (int)call.InArgs[1],
                            (int)call.InArgs[2]), null, 0, call.LogicalCallContext, call);
                    case "SetCellText":
                        SetCell(
                            (string)call.InArgs[0],
                            (int)call.InArgs[1],
                            (int)call.InArgs[2],
                            (string)call.InArgs[3]);
                        return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
                    case "ClearRange":
                        ClearRange(
                            (string)call.InArgs[0],
                            (int)call.InArgs[1],
                            (int)call.InArgs[2],
                            (int)call.InArgs[3],
                            (int)call.InArgs[4]);
                        return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
                    case "ClearWorksheet":
                        ClearWorksheet((string)call.InArgs[0]);
                        return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
                    case "MergeCells":
                        Merges.Add(new MergeRecord
                        {
                            SheetName = (string)call.InArgs[0],
                            Row = (int)call.InArgs[1],
                            Column = (int)call.InArgs[2],
                            RowSpan = (int)call.InArgs[3],
                            ColumnSpan = (int)call.InArgs[4],
                        });
                        return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
                    case "GetLastUsedRow":
                        return new ReturnMessage(GetLastUsedRow((string)call.InArgs[0]), null, 0, call.LogicalCallContext, call);
                    case "GetLastUsedColumn":
                        return new ReturnMessage(GetLastUsedColumn((string)call.InArgs[0]), null, 0, call.LogicalCallContext, call);
                    default:
                        throw new NotSupportedException(call.MethodName);
                }
            }

            public void SetCell(string sheetName, int row, int column, string value)
            {
                cells[BuildKey(sheetName, row, column)] = value ?? string.Empty;
            }

            public string GetCell(string sheetName, int row, int column)
            {
                return cells.TryGetValue(BuildKey(sheetName, row, column), out var value)
                    ? value
                    : string.Empty;
            }

            public new object GetTransparentProxy()
            {
                return base.GetTransparentProxy();
            }

            private void ClearRange(string sheetName, int startRow, int endRow, int startColumn, int endColumn)
            {
                ClearedRanges.Add(new ClearRangeRecord
                {
                    SheetName = sheetName,
                    StartRow = startRow,
                    EndRow = endRow,
                    StartColumn = startColumn,
                    EndColumn = endColumn,
                });

                var keysToRemove = cells.Keys
                    .Where(key => IsWithinRange(key, sheetName, startRow, endRow, startColumn, endColumn))
                    .ToArray();

                foreach (var key in keysToRemove)
                {
                    cells.Remove(key);
                }
            }

            private void ClearWorksheet(string sheetName)
            {
                var keysToRemove = cells.Keys
                    .Where(key => key.StartsWith(sheetName + "|", StringComparison.OrdinalIgnoreCase))
                    .ToArray();

                foreach (var key in keysToRemove)
                {
                    cells.Remove(key);
                }
            }

            private int GetLastUsedRow(string sheetName)
            {
                var prefix = sheetName + "|";
                var rows = cells.Keys
                    .Where(key => key.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    .Select(key => int.Parse(key.Split('|')[1]))
                    .ToArray();

                return rows.Length == 0 ? 0 : rows.Max();
            }

            private int GetLastUsedColumn(string sheetName)
            {
                var prefix = sheetName + "|";
                var columns = cells.Keys
                    .Where(key => key.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    .Select(key => int.Parse(key.Split('|')[2]))
                    .ToArray();

                return columns.Length == 0 ? 0 : columns.Max();
            }

            private static bool IsWithinRange(
                string key,
                string sheetName,
                int startRow,
                int endRow,
                int startColumn,
                int endColumn)
            {
                var parts = key.Split('|');
                if (parts.Length != 3)
                {
                    return false;
                }

                if (!string.Equals(parts[0], sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                var row = int.Parse(parts[1]);
                var column = int.Parse(parts[2]);
                return row >= startRow &&
                       row <= endRow &&
                       column >= startColumn &&
                       column <= endColumn;
            }

            private static string BuildKey(string sheetName, int row, int column)
            {
                return string.Join("|", sheetName ?? string.Empty, row, column);
            }
        }

        public sealed class MergeRecord
        {
            public string SheetName { get; set; }
            public int Row { get; set; }
            public int Column { get; set; }
            public int RowSpan { get; set; }
            public int ColumnSpan { get; set; }
        }

        public sealed class ClearRangeRecord
        {
            public string SheetName { get; set; }
            public int StartRow { get; set; }
            public int EndRow { get; set; }
            public int StartColumn { get; set; }
            public int EndColumn { get; set; }
        }
    }
}
