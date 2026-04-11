using System;
using System.Collections.Generic;
using System.Linq;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
using OfficeAgent.Core.Sync;
using OfficeAgent.ExcelAddIn.Excel;

namespace OfficeAgent.ExcelAddIn
{
    internal sealed class WorksheetDownloadPlan
    {
        public string OperationName { get; set; } = string.Empty;
        public string SheetName { get; set; } = string.Empty;
        public WorksheetSchema Schema { get; set; }
        public IReadOnlyList<IDictionary<string, object>> Rows { get; set; } = Array.Empty<IDictionary<string, object>>();
        public SyncOperationPreview Preview { get; set; }
        public ResolvedSelection Selection { get; set; }
    }

    internal sealed class WorksheetUploadPlan
    {
        public string OperationName { get; set; } = string.Empty;
        public string SheetName { get; set; } = string.Empty;
        public SyncOperationPreview Preview { get; set; } = new SyncOperationPreview();
        public bool ReplaceSnapshot { get; set; }
    }

    internal sealed class WorksheetSyncExecutionService
    {
        private const int HeaderRowCount = 2;
        private const int DataStartRow = 3;

        private readonly WorksheetSyncService worksheetSyncService;
        private readonly IWorksheetMetadataStore metadataStore;
        private readonly IWorksheetSelectionReader selectionReader;
        private readonly IWorksheetGridAdapter gridAdapter;
        private readonly WorksheetSelectionResolver selectionResolver;
        private readonly WorksheetSchemaLayoutService layoutService;
        private readonly SyncOperationPreviewFactory previewFactory;

        public WorksheetSyncExecutionService(
            WorksheetSyncService worksheetSyncService,
            IWorksheetMetadataStore metadataStore,
            IWorksheetSelectionReader selectionReader,
            IWorksheetGridAdapter gridAdapter,
            SyncOperationPreviewFactory previewFactory)
        {
            this.worksheetSyncService = worksheetSyncService ?? throw new ArgumentNullException(nameof(worksheetSyncService));
            this.metadataStore = metadataStore ?? throw new ArgumentNullException(nameof(metadataStore));
            this.selectionReader = selectionReader ?? throw new ArgumentNullException(nameof(selectionReader));
            this.gridAdapter = gridAdapter ?? throw new ArgumentNullException(nameof(gridAdapter));
            this.previewFactory = previewFactory ?? throw new ArgumentNullException(nameof(previewFactory));
            selectionResolver = new WorksheetSelectionResolver();
            layoutService = new WorksheetSchemaLayoutService();
        }

        public WorksheetDownloadPlan PrepareFullDownload(string sheetName)
        {
            var schema = worksheetSyncService.LoadSchemaForSheet(sheetName);
            var rows = worksheetSyncService.ExecutePartialDownload(sheetName, new ResolvedSelection());
            var overwritePreview = worksheetSyncService.PrepareIncrementalUpload(sheetName, ReadAllCurrentCells(sheetName, schema));

            return new WorksheetDownloadPlan
            {
                OperationName = "全量下载",
                SheetName = sheetName,
                Schema = schema,
                Rows = rows,
                Preview = overwritePreview,
            };
        }

        public WorksheetDownloadPlan PreparePartialDownload(string sheetName)
        {
            var schema = worksheetSyncService.LoadSchemaForSheet(sheetName);
            var selection = ResolveCurrentSelection(sheetName, schema);
            var rows = selection.RowIds.Length == 0
                ? Array.Empty<IDictionary<string, object>>()
                : worksheetSyncService.ExecutePartialDownload(sheetName, selection);
            var overwritePreview = worksheetSyncService.PrepareIncrementalUpload(sheetName, ReadSelectionChanges(sheetName, schema, selection));

            return new WorksheetDownloadPlan
            {
                OperationName = "部分下载",
                SheetName = sheetName,
                Schema = schema,
                Rows = rows,
                Preview = overwritePreview,
                Selection = selection,
            };
        }

        public void ExecuteDownload(WorksheetDownloadPlan plan)
        {
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }

            if (plan.Selection == null)
            {
                WriteFullWorksheet(plan);
                metadataStore.SaveSnapshot(
                    plan.SheetName,
                    BuildSnapshotCellsFromRows(plan.SheetName, plan.Schema, plan.Rows, includedFieldKeys: null));
                return;
            }

            WritePartialCells(plan);
            MergeSnapshot(
                plan.SheetName,
                BuildSnapshotCellsFromRows(plan.SheetName, plan.Schema, plan.Rows, plan.Selection.ApiFieldKeys));
        }

        public WorksheetUploadPlan PrepareFullUpload(string sheetName)
        {
            var schema = worksheetSyncService.LoadSchemaForSheet(sheetName);
            var changes = ReadAllCurrentCells(sheetName, schema);

            return new WorksheetUploadPlan
            {
                OperationName = "全量上传",
                SheetName = sheetName,
                Preview = BuildUploadPreview("全量上传", changes),
                ReplaceSnapshot = true,
            };
        }

        public WorksheetUploadPlan PreparePartialUpload(string sheetName)
        {
            var schema = worksheetSyncService.LoadSchemaForSheet(sheetName);
            var selection = ResolveCurrentSelection(sheetName, schema);
            var changes = ReadSelectionChanges(sheetName, schema, selection);

            return new WorksheetUploadPlan
            {
                OperationName = "部分上传",
                SheetName = sheetName,
                Preview = BuildUploadPreview("部分上传", changes),
                ReplaceSnapshot = false,
            };
        }

        public WorksheetUploadPlan PrepareIncrementalUpload(string sheetName)
        {
            var schema = worksheetSyncService.LoadSchemaForSheet(sheetName);
            var currentCells = ReadAllCurrentCells(sheetName, schema);
            var preview = worksheetSyncService.PrepareIncrementalUpload(sheetName, currentCells);
            preview.OperationName = "增量上传";
            preview.Summary = $"增量上传将提交 {preview.Changes.Length} 个单元格。";

            return new WorksheetUploadPlan
            {
                OperationName = "增量上传",
                SheetName = sheetName,
                Preview = preview,
                ReplaceSnapshot = false,
            };
        }

        public void ExecuteUpload(WorksheetUploadPlan plan)
        {
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }

            var changes = plan.Preview?.Changes ?? Array.Empty<CellChange>();
            if (changes.Length == 0)
            {
                return;
            }

            worksheetSyncService.ExecutePartialUpload(plan.SheetName, changes);
            var snapshotCells = BuildSnapshotCellsFromChanges(plan.SheetName, changes);

            if (plan.ReplaceSnapshot)
            {
                metadataStore.SaveSnapshot(plan.SheetName, snapshotCells);
                return;
            }

            MergeSnapshot(plan.SheetName, snapshotCells);
        }

        private void WriteFullWorksheet(WorksheetDownloadPlan plan)
        {
            gridAdapter.ClearWorksheet(plan.SheetName);

            var headerPlan = layoutService.BuildHeaderPlan(plan.Schema);
            foreach (var headerCell in headerPlan)
            {
                gridAdapter.SetCellText(plan.SheetName, headerCell.Row, headerCell.Column, headerCell.Text);
                gridAdapter.MergeCells(plan.SheetName, headerCell.Row, headerCell.Column, headerCell.RowSpan, headerCell.ColumnSpan);
            }

            for (var rowIndex = 0; rowIndex < plan.Rows.Count; rowIndex++)
            {
                var row = plan.Rows[rowIndex];
                var targetRow = DataStartRow + rowIndex;
                foreach (var column in plan.Schema.Columns ?? Array.Empty<WorksheetColumnBinding>())
                {
                    var value = GetRowValue(row, column.ApiFieldKey);
                    gridAdapter.SetCellText(plan.SheetName, targetRow, column.ColumnIndex, value);
                }
            }
        }

        private void WritePartialCells(WorksheetDownloadPlan plan)
        {
            var columnsByIndex = (plan.Schema.Columns ?? Array.Empty<WorksheetColumnBinding>())
                .ToDictionary(column => column.ColumnIndex, column => column);
            var rowsById = plan.Rows
                .Where(row => !string.IsNullOrWhiteSpace(GetRowId(plan.Schema, row)))
                .ToDictionary(row => GetRowId(plan.Schema, row), row => row, StringComparer.Ordinal);

            foreach (var targetCell in plan.Selection.TargetCells ?? Array.Empty<SelectedVisibleCell>())
            {
                if (!columnsByIndex.TryGetValue(targetCell.Column, out var column))
                {
                    continue;
                }

                var rowId = GetRowId(plan.SheetName, plan.Schema, targetCell.Row);
                if (string.IsNullOrWhiteSpace(rowId))
                {
                    continue;
                }

                if (!rowsById.TryGetValue(rowId, out var row))
                {
                    continue;
                }

                var value = GetRowValue(row, column.ApiFieldKey);
                gridAdapter.SetCellText(plan.SheetName, targetCell.Row, targetCell.Column, value);
            }
        }

        private ResolvedSelection ResolveCurrentSelection(string sheetName, WorksheetSchema schema)
        {
            var visibleCells = selectionReader.ReadVisibleSelection() ?? Array.Empty<SelectedVisibleCell>();
            return selectionResolver.Resolve(schema, visibleCells, row => GetRowId(sheetName, schema, row));
        }

        private CellChange[] ReadAllCurrentCells(string sheetName, WorksheetSchema schema)
        {
            var idColumn = GetIdColumn(schema);
            if (idColumn == null)
            {
                return Array.Empty<CellChange>();
            }

            var snapshotLookup = BuildSnapshotLookup(sheetName);
            var lastUsedRow = gridAdapter.GetLastUsedRow(sheetName);
            var result = new List<CellChange>();

            for (var row = DataStartRow; row <= lastUsedRow; row++)
            {
                var rowId = gridAdapter.GetCellText(sheetName, row, idColumn.ColumnIndex);
                if (string.IsNullOrWhiteSpace(rowId))
                {
                    continue;
                }

                foreach (var column in schema.Columns.Where(item => !item.IsIdColumn))
                {
                    var key = BuildSnapshotKey(rowId, column.ApiFieldKey);
                    snapshotLookup.TryGetValue(key, out var oldValue);
                    result.Add(new CellChange
                    {
                        SheetName = sheetName,
                        RowId = rowId,
                        ApiFieldKey = column.ApiFieldKey,
                        OldValue = oldValue ?? string.Empty,
                        NewValue = gridAdapter.GetCellText(sheetName, row, column.ColumnIndex),
                    });
                }
            }

            return result.ToArray();
        }

        private CellChange[] ReadSelectionChanges(string sheetName, WorksheetSchema schema, ResolvedSelection selection)
        {
            var columnsByIndex = (schema.Columns ?? Array.Empty<WorksheetColumnBinding>())
                .ToDictionary(column => column.ColumnIndex, column => column);
            var snapshotLookup = BuildSnapshotLookup(sheetName);
            var result = new List<CellChange>();

            foreach (var targetCell in selection.TargetCells ?? Array.Empty<SelectedVisibleCell>())
            {
                if (!columnsByIndex.TryGetValue(targetCell.Column, out var column) || column.IsIdColumn)
                {
                    continue;
                }

                var rowId = GetRowId(sheetName, schema, targetCell.Row);
                if (string.IsNullOrWhiteSpace(rowId))
                {
                    continue;
                }

                var key = BuildSnapshotKey(rowId, column.ApiFieldKey);
                snapshotLookup.TryGetValue(key, out var oldValue);
                result.Add(new CellChange
                {
                    SheetName = sheetName,
                    RowId = rowId,
                    ApiFieldKey = column.ApiFieldKey,
                    OldValue = oldValue ?? string.Empty,
                    NewValue = gridAdapter.GetCellText(sheetName, targetCell.Row, targetCell.Column),
                });
            }

            return result.ToArray();
        }

        private WorksheetSnapshotCell[] BuildSnapshotCellsFromRows(
            string sheetName,
            WorksheetSchema schema,
            IReadOnlyList<IDictionary<string, object>> rows,
            IReadOnlyList<string> includedFieldKeys)
        {
            var allowedFieldKeys = includedFieldKeys == null
                ? null
                : new HashSet<string>(includedFieldKeys, StringComparer.Ordinal);
            var result = new List<WorksheetSnapshotCell>();

            foreach (var row in rows ?? Array.Empty<IDictionary<string, object>>())
            {
                var rowId = GetRowId(schema, row);
                if (string.IsNullOrWhiteSpace(rowId))
                {
                    continue;
                }

                foreach (var column in schema.Columns.Where(item => !item.IsIdColumn))
                {
                    if (allowedFieldKeys != null && !allowedFieldKeys.Contains(column.ApiFieldKey))
                    {
                        continue;
                    }

                    result.Add(new WorksheetSnapshotCell
                    {
                        SheetName = sheetName,
                        RowId = rowId,
                        ApiFieldKey = column.ApiFieldKey,
                        Value = GetRowValue(row, column.ApiFieldKey),
                    });
                }
            }

            return result.ToArray();
        }

        private static WorksheetSnapshotCell[] BuildSnapshotCellsFromChanges(string sheetName, IReadOnlyList<CellChange> changes)
        {
            return (changes ?? Array.Empty<CellChange>())
                .Select(change => new WorksheetSnapshotCell
                {
                    SheetName = sheetName,
                    RowId = change.RowId,
                    ApiFieldKey = change.ApiFieldKey,
                    Value = change.NewValue ?? string.Empty,
                })
                .ToArray();
        }

        private void MergeSnapshot(string sheetName, IReadOnlyList<WorksheetSnapshotCell> updates)
        {
            var merged = (metadataStore.LoadSnapshot(sheetName) ?? Array.Empty<WorksheetSnapshotCell>())
                .ToDictionary(
                    item => BuildSnapshotKey(item.RowId, item.ApiFieldKey),
                    item => new WorksheetSnapshotCell
                    {
                        SheetName = sheetName,
                        RowId = item.RowId,
                        ApiFieldKey = item.ApiFieldKey,
                        Value = item.Value,
                    },
                    StringComparer.Ordinal);

            foreach (var update in updates ?? Array.Empty<WorksheetSnapshotCell>())
            {
                merged[BuildSnapshotKey(update.RowId, update.ApiFieldKey)] = new WorksheetSnapshotCell
                {
                    SheetName = sheetName,
                    RowId = update.RowId,
                    ApiFieldKey = update.ApiFieldKey,
                    Value = update.Value,
                };
            }

            metadataStore.SaveSnapshot(sheetName, merged.Values.ToArray());
        }

        private Dictionary<string, string> BuildSnapshotLookup(string sheetName)
        {
            return (metadataStore.LoadSnapshot(sheetName) ?? Array.Empty<WorksheetSnapshotCell>())
                .ToDictionary(
                    item => BuildSnapshotKey(item.RowId, item.ApiFieldKey),
                    item => item.Value ?? string.Empty,
                    StringComparer.Ordinal);
        }

        private SyncOperationPreview BuildUploadPreview(string operationName, IReadOnlyList<CellChange> changes)
        {
            var preview = previewFactory.CreateUploadPreview(operationName, changes);
            preview.OperationName = operationName;
            preview.Summary = $"{operationName}将提交 {preview.Changes.Length} 个单元格。";
            return preview;
        }

        private string GetRowId(string sheetName, WorksheetSchema schema, int row)
        {
            var idColumn = GetIdColumn(schema);
            if (idColumn == null)
            {
                return string.Empty;
            }

            return gridAdapter.GetCellText(sheetName, row, idColumn.ColumnIndex);
        }

        private static WorksheetColumnBinding GetIdColumn(WorksheetSchema schema)
        {
            return (schema?.Columns ?? Array.Empty<WorksheetColumnBinding>())
                .FirstOrDefault(column => column.IsIdColumn);
        }

        private static string GetRowId(WorksheetSchema schema, IDictionary<string, object> row)
        {
            var idColumn = GetIdColumn(schema);
            return idColumn == null ? string.Empty : GetRowValue(row, idColumn.ApiFieldKey);
        }

        private static string GetRowValue(IDictionary<string, object> row, string fieldKey)
        {
            if (row == null || string.IsNullOrWhiteSpace(fieldKey))
            {
                return string.Empty;
            }

            if (row.TryGetValue(fieldKey, out var value))
            {
                return Convert.ToString(value) ?? string.Empty;
            }

            foreach (var item in row)
            {
                if (string.Equals(item.Key, fieldKey, StringComparison.OrdinalIgnoreCase))
                {
                    return Convert.ToString(item.Value) ?? string.Empty;
                }
            }

            return string.Empty;
        }

        private static string BuildSnapshotKey(string rowId, string apiFieldKey)
        {
            return string.Concat(rowId ?? string.Empty, "|", apiFieldKey ?? string.Empty);
        }
    }
}
