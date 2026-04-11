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
        public void ExecuteFullDownloadClearsSheetWritesHeadersRowsAndSnapshot()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            metadataStore.Bindings["Sync-performance"] = new SheetBinding
            {
                SheetName = "Sync-performance",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
            };

            connector.SchemaByProjectId["performance"] = CreateSchema();
            connector.FindResult = new[]
            {
                CreateRow("row-1", "项目 A", "2026-01-02", "2026-01-05"),
            };

            var selectionReader = new FakeWorksheetSelectionReader();
            var (service, grid) = CreateService(connector, metadataStore, selectionReader);

            var plan = InvokePrepare(service, "PrepareFullDownload", "Sync-performance");
            InvokeExecute(service, "ExecuteDownload", plan);

            Assert.Contains("Sync-performance", grid.ClearedSheets);
            Assert.Equal("行 ID", grid.GetCell("Sync-performance", 1, 1));
            Assert.Equal("名称", grid.GetCell("Sync-performance", 1, 2));
            Assert.Equal("测试活动111", grid.GetCell("Sync-performance", 1, 3));
            Assert.Equal("开始时间", grid.GetCell("Sync-performance", 2, 3));
            Assert.Equal("结束时间", grid.GetCell("Sync-performance", 2, 4));
            Assert.Equal("row-1", grid.GetCell("Sync-performance", 3, 1));
            Assert.Equal("项目 A", grid.GetCell("Sync-performance", 3, 2));
            Assert.Equal("2026-01-02", grid.GetCell("Sync-performance", 3, 3));
            Assert.Equal("2026-01-05", grid.GetCell("Sync-performance", 3, 4));

            Assert.Contains(grid.Merges, merge => merge.SheetName == "Sync-performance" && merge.Row == 1 && merge.Column == 1 && merge.RowSpan == 2 && merge.ColumnSpan == 1);
            Assert.Contains(grid.Merges, merge => merge.SheetName == "Sync-performance" && merge.Row == 1 && merge.Column == 3 && merge.RowSpan == 1 && merge.ColumnSpan == 2);

            Assert.NotNull(metadataStore.LastSavedSnapshot);
            Assert.Equal(3, metadataStore.LastSavedSnapshot.Length);
            Assert.Contains(metadataStore.LastSavedSnapshot, cell => cell.RowId == "row-1" && cell.ApiFieldKey == "name" && cell.Value == "项目 A");
            Assert.Contains(metadataStore.LastSavedSnapshot, cell => cell.RowId == "row-1" && cell.ApiFieldKey == "start_12345678" && cell.Value == "2026-01-02");
            Assert.Contains(metadataStore.LastSavedSnapshot, cell => cell.RowId == "row-1" && cell.ApiFieldKey == "end_12345678" && cell.Value == "2026-01-05");
        }

        [Fact]
        public void ExecutePartialDownloadWritesOnlySelectedVisibleCellsAndMergesSnapshot()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            metadataStore.Bindings["Sync-performance"] = new SheetBinding
            {
                SheetName = "Sync-performance",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
            };
            metadataStore.StoredSnapshot = new[]
            {
                new WorksheetSnapshotCell { SheetName = "Sync-performance", RowId = "row-1", ApiFieldKey = "name", Value = "项目 A" },
            };

            connector.SchemaByProjectId["performance"] = CreateSchema();
            connector.FindResult = new[]
            {
                CreateRow("row-1", "项目 A", "2026-02-01", "2026-02-08"),
            };

            var selectionReader = new FakeWorksheetSelectionReader
            {
                VisibleCells = new[]
                {
                    new SelectedVisibleCell { Row = 3, Column = 3, Value = "旧开始时间" },
                },
            };
            var (service, grid) = CreateService(connector, metadataStore, selectionReader);
            grid.SetCell("Sync-performance", 3, 1, "row-1");
            grid.SetCell("Sync-performance", 3, 2, "项目 A");
            grid.SetCell("Sync-performance", 3, 3, "旧开始时间");
            grid.SetCell("Sync-performance", 3, 4, "旧结束时间");

            var plan = InvokePrepare(service, "PreparePartialDownload", "Sync-performance");
            InvokeExecute(service, "ExecuteDownload", plan);

            Assert.Equal("2026-02-01", grid.GetCell("Sync-performance", 3, 3));
            Assert.Equal("旧结束时间", grid.GetCell("Sync-performance", 3, 4));

            Assert.NotNull(metadataStore.LastSavedSnapshot);
            Assert.Equal(2, metadataStore.LastSavedSnapshot.Length);
            Assert.Contains(metadataStore.LastSavedSnapshot, cell => cell.RowId == "row-1" && cell.ApiFieldKey == "name" && cell.Value == "项目 A");
            Assert.Contains(metadataStore.LastSavedSnapshot, cell => cell.RowId == "row-1" && cell.ApiFieldKey == "start_12345678" && cell.Value == "2026-02-01");
        }

        [Fact]
        public void ExecuteFullUploadSendsAllNonIdCellsForRowsWithIdsAndRefreshesSnapshot()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            metadataStore.Bindings["Sync-performance"] = new SheetBinding
            {
                SheetName = "Sync-performance",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
            };
            connector.SchemaByProjectId["performance"] = CreateSchema();

            var selectionReader = new FakeWorksheetSelectionReader();
            var (service, grid) = CreateService(connector, metadataStore, selectionReader);
            grid.SetCell("Sync-performance", 3, 1, "row-1");
            grid.SetCell("Sync-performance", 3, 2, "项目 A");
            grid.SetCell("Sync-performance", 3, 3, "2026-01-02");
            grid.SetCell("Sync-performance", 3, 4, "2026-01-05");
            grid.SetCell("Sync-performance", 4, 1, string.Empty);
            grid.SetCell("Sync-performance", 4, 2, "无 ID 行");
            grid.SetCell("Sync-performance", 4, 3, "2026-03-01");
            grid.SetCell("Sync-performance", 4, 4, "2026-03-05");

            var plan = InvokePrepare(service, "PrepareFullUpload", "Sync-performance");
            var preview = ReadPreview(plan);
            Assert.Equal(3, preview.Changes.Length);

            InvokeExecute(service, "ExecuteUpload", plan);

            Assert.Equal("performance", connector.LastBatchSaveProjectId);
            Assert.Equal(3, connector.LastBatchSaveChanges.Count);
            Assert.Contains(connector.LastBatchSaveChanges, change => change.RowId == "row-1" && change.ApiFieldKey == "name" && change.NewValue == "项目 A");
            Assert.Contains(connector.LastBatchSaveChanges, change => change.RowId == "row-1" && change.ApiFieldKey == "start_12345678");
            Assert.DoesNotContain(connector.LastBatchSaveChanges, change => string.IsNullOrWhiteSpace(change.RowId));

            Assert.NotNull(metadataStore.LastSavedSnapshot);
            Assert.Equal(3, metadataStore.LastSavedSnapshot.Length);
        }

        [Fact]
        public void ExecuteIncrementalUploadSendsOnlyDirtySnapshotBackedCells()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            metadataStore.Bindings["Sync-performance"] = new SheetBinding
            {
                SheetName = "Sync-performance",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
            };
            metadataStore.StoredSnapshot = new[]
            {
                new WorksheetSnapshotCell { SheetName = "Sync-performance", RowId = "row-1", ApiFieldKey = "name", Value = "旧项目名" },
                new WorksheetSnapshotCell { SheetName = "Sync-performance", RowId = "row-1", ApiFieldKey = "start_12345678", Value = "2026-01-02" },
            };
            connector.SchemaByProjectId["performance"] = CreateSchema();

            var selectionReader = new FakeWorksheetSelectionReader();
            var (service, grid) = CreateService(connector, metadataStore, selectionReader);
            grid.SetCell("Sync-performance", 3, 1, "row-1");
            grid.SetCell("Sync-performance", 3, 2, "新项目名");
            grid.SetCell("Sync-performance", 3, 3, "2026-01-02");
            grid.SetCell("Sync-performance", 3, 4, string.Empty);
            grid.SetCell("Sync-performance", 4, 1, "row-new");
            grid.SetCell("Sync-performance", 4, 2, "新行");
            grid.SetCell("Sync-performance", 4, 3, "2026-04-01");
            grid.SetCell("Sync-performance", 4, 4, "2026-04-08");

            var plan = InvokePrepare(service, "PrepareIncrementalUpload", "Sync-performance");
            var preview = ReadPreview(plan);
            Assert.Single(preview.Changes);
            Assert.Equal("row-1", preview.Changes[0].RowId);
            Assert.Equal("name", preview.Changes[0].ApiFieldKey);

            InvokeExecute(service, "ExecuteUpload", plan);

            Assert.Equal("performance", connector.LastBatchSaveProjectId);
            Assert.Single(connector.LastBatchSaveChanges);
            Assert.Equal("新项目名", connector.LastBatchSaveChanges[0].NewValue);

            Assert.NotNull(metadataStore.LastSavedSnapshot);
            Assert.Contains(metadataStore.LastSavedSnapshot, cell => cell.RowId == "row-1" && cell.ApiFieldKey == "name" && cell.Value == "新项目名");
            Assert.Contains(metadataStore.LastSavedSnapshot, cell => cell.RowId == "row-1" && cell.ApiFieldKey == "start_12345678" && cell.Value == "2026-01-02");
            Assert.DoesNotContain(metadataStore.LastSavedSnapshot, cell => cell.RowId == "row-new");
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

        private static WorksheetSchema CreateSchema()
        {
            return new WorksheetSchema
            {
                SystemKey = "current-business-system",
                ProjectId = "performance",
                Columns = new[]
                {
                    new WorksheetColumnBinding
                    {
                        ColumnIndex = 1,
                        ApiFieldKey = "id",
                        ColumnKind = WorksheetColumnKind.Single,
                        ParentHeaderText = "行 ID",
                        ChildHeaderText = "行 ID",
                        IsIdColumn = true,
                    },
                    new WorksheetColumnBinding
                    {
                        ColumnIndex = 2,
                        ApiFieldKey = "name",
                        ColumnKind = WorksheetColumnKind.Single,
                        ParentHeaderText = "名称",
                        ChildHeaderText = "名称",
                    },
                    new WorksheetColumnBinding
                    {
                        ColumnIndex = 3,
                        ApiFieldKey = "start_12345678",
                        ColumnKind = WorksheetColumnKind.ActivityProperty,
                        ParentHeaderText = "测试活动111",
                        ChildHeaderText = "开始时间",
                        ActivityId = "12345678",
                        ActivityName = "测试活动111",
                        PropertyKey = "start",
                    },
                    new WorksheetColumnBinding
                    {
                        ColumnIndex = 4,
                        ApiFieldKey = "end_12345678",
                        ColumnKind = WorksheetColumnKind.ActivityProperty,
                        ParentHeaderText = "测试活动111",
                        ChildHeaderText = "结束时间",
                        ActivityId = "12345678",
                        ActivityName = "测试活动111",
                        PropertyKey = "end",
                    },
                },
            };
        }

        private static IDictionary<string, object> CreateRow(string id, string name, string start, string end)
        {
            return new Dictionary<string, object>(StringComparer.Ordinal)
            {
                ["id"] = id,
                ["name"] = name,
                ["start_12345678"] = start,
                ["end_12345678"] = end,
            };
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

        private sealed class FakeSystemConnector : ISystemConnector
        {
            public Dictionary<string, WorksheetSchema> SchemaByProjectId { get; } = new Dictionary<string, WorksheetSchema>(StringComparer.Ordinal);

            public IReadOnlyList<IDictionary<string, object>> FindResult { get; set; } = Array.Empty<IDictionary<string, object>>();

            public string LastBatchSaveProjectId { get; private set; }

            public IReadOnlyList<CellChange> LastBatchSaveChanges { get; private set; } = Array.Empty<CellChange>();

            public IReadOnlyList<ProjectOption> GetProjects()
            {
                return Array.Empty<ProjectOption>();
            }

            public WorksheetSchema GetSchema(string projectId)
            {
                return SchemaByProjectId[projectId];
            }

            public IReadOnlyList<IDictionary<string, object>> Find(string projectId, IReadOnlyList<string> rowIds, IReadOnlyList<string> fieldKeys)
            {
                IEnumerable<IDictionary<string, object>> rows = FindResult;

                if (rowIds != null && rowIds.Count > 0)
                {
                    rows = rows.Where(row => rowIds.Contains(row["id"]?.ToString()));
                }

                if (fieldKeys != null && fieldKeys.Count > 0)
                {
                    rows = rows.Select(row =>
                    {
                        var filtered = new Dictionary<string, object>(StringComparer.Ordinal)
                        {
                            ["id"] = row["id"],
                        };

                        foreach (var fieldKey in fieldKeys)
                        {
                            if (row.TryGetValue(fieldKey, out var value))
                            {
                                filtered[fieldKey] = value;
                            }
                        }

                        return (IDictionary<string, object>)filtered;
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

            public WorksheetSnapshotCell[] StoredSnapshot { get; set; } = Array.Empty<WorksheetSnapshotCell>();

            public WorksheetSnapshotCell[] LastSavedSnapshot { get; private set; }

            public void SaveBinding(SheetBinding binding)
            {
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

            public WorksheetSnapshotCell[] LoadSnapshot(string sheetName)
            {
                return StoredSnapshot
                    .Where(cell => string.Equals(cell.SheetName, sheetName, StringComparison.OrdinalIgnoreCase))
                    .ToArray();
            }

            public void SaveSnapshot(string sheetName, WorksheetSnapshotCell[] cells)
            {
                LastSavedSnapshot = cells?.ToArray() ?? Array.Empty<WorksheetSnapshotCell>();
                StoredSnapshot = LastSavedSnapshot
                    .Select(cell => new WorksheetSnapshotCell
                    {
                        SheetName = sheetName,
                        RowId = cell.RowId,
                        ApiFieldKey = cell.ApiFieldKey,
                        Value = cell.Value,
                    })
                    .ToArray();
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

            public List<string> ClearedSheets { get; } = new List<string>();

            public List<MergeRecord> Merges { get; } = new List<MergeRecord>();

            public FakeWorksheetGridAdapter(Type interfaceType)
                : base(interfaceType)
            {
            }

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
                    case "ClearWorksheet":
                        ClearedSheets.Add((string)call.InArgs[0]);
                        ClearSheet((string)call.InArgs[0]);
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

            private void ClearSheet(string sheetName)
            {
                var keys = cells.Keys
                    .Where(key => key.StartsWith(sheetName + "|", StringComparison.OrdinalIgnoreCase))
                    .ToArray();

                foreach (var key in keys)
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
    }
}
