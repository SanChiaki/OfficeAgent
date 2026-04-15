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
        public SheetBinding Binding { get; set; }
        public WorksheetSchema Schema { get; set; }
        public WorksheetRuntimeColumn[] RuntimeColumns { get; set; } = Array.Empty<WorksheetRuntimeColumn>();
        public IReadOnlyList<IDictionary<string, object>> Rows { get; set; } = Array.Empty<IDictionary<string, object>>();
        public SyncOperationPreview Preview { get; set; } = new SyncOperationPreview();
        public ResolvedSelection Selection { get; set; }
        public bool UsesExistingLayout { get; set; }
    }

    internal sealed class WorksheetUploadPlan
    {
        public string OperationName { get; set; } = string.Empty;
        public string SheetName { get; set; } = string.Empty;
        public string SystemKey { get; set; } = string.Empty;
        public string ProjectId { get; set; } = string.Empty;
        public SyncOperationPreview Preview { get; set; } = new SyncOperationPreview();
    }

    internal sealed class WorksheetSyncExecutionService
    {
        private readonly WorksheetSyncService worksheetSyncService;
        private readonly IWorksheetSelectionReader selectionReader;
        private readonly IWorksheetGridAdapter gridAdapter;
        private readonly WorksheetSelectionResolver selectionResolver;
        private readonly WorksheetSchemaLayoutService layoutService;
        private readonly SyncOperationPreviewFactory previewFactory;
        private readonly WorksheetHeaderMatcher headerMatcher;
        private readonly FieldMappingValueAccessor valueAccessor;

        public WorksheetSyncExecutionService(
            WorksheetSyncService worksheetSyncService,
            IWorksheetMetadataStore metadataStore,
            IWorksheetSelectionReader selectionReader,
            IWorksheetGridAdapter gridAdapter,
            SyncOperationPreviewFactory previewFactory)
        {
            this.worksheetSyncService = worksheetSyncService ?? throw new ArgumentNullException(nameof(worksheetSyncService));
            _ = metadataStore ?? throw new ArgumentNullException(nameof(metadataStore));
            this.selectionReader = selectionReader ?? throw new ArgumentNullException(nameof(selectionReader));
            this.gridAdapter = gridAdapter ?? throw new ArgumentNullException(nameof(gridAdapter));
            this.previewFactory = previewFactory ?? throw new ArgumentNullException(nameof(previewFactory));
            selectionResolver = new WorksheetSelectionResolver();
            layoutService = new WorksheetSchemaLayoutService();
            valueAccessor = new FieldMappingValueAccessor();
            headerMatcher = new WorksheetHeaderMatcher(valueAccessor);
        }

        public void InitializeCurrentSheet(string sheetName, ProjectOption project)
        {
            worksheetSyncService.InitializeSheet(sheetName, project);
        }

        public void TryAutoInitializeCurrentSheet(string sheetName, ProjectOption project)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                throw new ArgumentException("Sheet name is required.", nameof(sheetName));
            }

            if (project == null)
            {
                throw new ArgumentNullException(nameof(project));
            }

            try
            {
                var binding = worksheetSyncService.LoadBinding(sheetName);
                if (!string.Equals(binding.SystemKey, project.SystemKey, StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(binding.ProjectId, project.ProjectId, StringComparison.OrdinalIgnoreCase))
                {
                    InitializeCurrentSheet(sheetName, project);
                    return;
                }

                var definition = worksheetSyncService.LoadFieldMappingDefinition(binding.SystemKey, binding.ProjectId);
                var mappings = worksheetSyncService.LoadFieldMappings(sheetName, binding.SystemKey, binding.ProjectId);
                if (!HasUsableMappings(definition, mappings))
                {
                    InitializeCurrentSheet(sheetName, project);
                }
            }
            catch (InvalidOperationException)
            {
                InitializeCurrentSheet(sheetName, project);
            }
        }

        public WorksheetDownloadPlan PrepareFullDownload(string sheetName)
        {
            var context = ResolveFullDownloadContext(sheetName);
            var rows = worksheetSyncService.Download(
                context.Binding.SystemKey,
                context.Binding.ProjectId,
                Array.Empty<string>(),
                GetRequestedFieldKeys(context.RuntimeColumns));

            return new WorksheetDownloadPlan
            {
                OperationName = "全量下载",
                SheetName = sheetName,
                Binding = context.Binding,
                Schema = context.Schema,
                RuntimeColumns = context.RuntimeColumns,
                Rows = rows,
                Preview = new SyncOperationPreview { OperationName = "全量下载" },
                UsesExistingLayout = context.UsesExistingLayout,
            };
        }

        public WorksheetDownloadPlan PreparePartialDownload(string sheetName)
        {
            var context = ResolveMatchedSheetContext(sheetName);
            var selection = ResolveCurrentSelection(sheetName, context.Schema);
            var rows = selection.RowIds.Length == 0
                ? Array.Empty<IDictionary<string, object>>()
                : worksheetSyncService.Download(context.Binding.SystemKey, context.Binding.ProjectId, selection.RowIds, selection.ApiFieldKeys);

            return new WorksheetDownloadPlan
            {
                OperationName = "部分下载",
                SheetName = sheetName,
                Binding = context.Binding,
                Schema = context.Schema,
                RuntimeColumns = context.RuntimeColumns,
                Rows = rows,
                Preview = new SyncOperationPreview { OperationName = "部分下载" },
                Selection = selection,
                UsesExistingLayout = true,
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
                return;
            }

            WritePartialCells(plan);
        }

        public WorksheetUploadPlan PrepareFullUpload(string sheetName)
        {
            var context = ResolveMatchedSheetContext(sheetName);
            var changes = ReadAllCurrentCells(sheetName, context.Binding, context.Schema);

            return new WorksheetUploadPlan
            {
                OperationName = "全量上传",
                SheetName = sheetName,
                ProjectId = context.Binding.ProjectId,
                SystemKey = context.Binding.SystemKey,
                Preview = BuildUploadPreview("全量上传", changes),
            };
        }

        public WorksheetUploadPlan PreparePartialUpload(string sheetName)
        {
            var context = ResolveMatchedSheetContext(sheetName);
            var selection = ResolveCurrentSelection(sheetName, context.Schema);
            var changes = ReadSelectionChanges(sheetName, context.Schema, selection);

            return new WorksheetUploadPlan
            {
                OperationName = "部分上传",
                SheetName = sheetName,
                ProjectId = context.Binding.ProjectId,
                SystemKey = context.Binding.SystemKey,
                Preview = BuildUploadPreview("部分上传", changes),
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

            worksheetSyncService.Upload(plan.SystemKey, plan.ProjectId, changes);
        }

        private SheetExecutionContext ResolveFullDownloadContext(string sheetName)
        {
            var context = LoadSheetContext(sheetName);
            var runtimeColumns = LoadRuntimeColumns(sheetName, context.Binding, context.Definition, context.Mappings);

            if (runtimeColumns.Length > 0)
            {
                EnsureIdColumn(runtimeColumns);
                context.RuntimeColumns = runtimeColumns;
                context.Schema = BuildSchema(context.Binding, runtimeColumns);
                context.UsesExistingLayout = true;
                return context;
            }

            if (HasAnyHeaderText(sheetName, context.Binding))
            {
                throw CreateHeaderMatchException();
            }

            var configuredColumns = BuildConfiguredColumns(context.Binding, context.Definition, context.Mappings);
            if (configuredColumns.Length == 0)
            {
                throw CreateInitializationRequiredException();
            }

            EnsureIdColumn(configuredColumns);
            context.RuntimeColumns = configuredColumns;
            context.Schema = BuildSchema(context.Binding, configuredColumns);
            context.UsesExistingLayout = false;
            return context;
        }

        private SheetExecutionContext ResolveMatchedSheetContext(string sheetName)
        {
            var context = LoadSheetContext(sheetName);
            var runtimeColumns = LoadRuntimeColumns(sheetName, context.Binding, context.Definition, context.Mappings);

            if (runtimeColumns.Length == 0)
            {
                throw CreateHeaderMatchException();
            }

            EnsureIdColumn(runtimeColumns);
            context.RuntimeColumns = runtimeColumns;
            context.Schema = BuildSchema(context.Binding, runtimeColumns);
            context.UsesExistingLayout = true;
            return context;
        }

        private SheetExecutionContext LoadSheetContext(string sheetName)
        {
            var binding = worksheetSyncService.LoadBinding(sheetName);
            ValidateBinding(binding);
            var definition = worksheetSyncService.LoadFieldMappingDefinition(binding.SystemKey, binding.ProjectId);
            var mappings = worksheetSyncService.LoadFieldMappings(sheetName, binding.SystemKey, binding.ProjectId) ?? Array.Empty<SheetFieldMappingRow>();

            if (!HasUsableMappings(definition, mappings))
            {
                worksheetSyncService.InitializeSheet(sheetName, CreateProjectOption(binding));
                definition = worksheetSyncService.LoadFieldMappingDefinition(binding.SystemKey, binding.ProjectId);
                mappings = worksheetSyncService.LoadFieldMappings(sheetName, binding.SystemKey, binding.ProjectId) ?? Array.Empty<SheetFieldMappingRow>();
            }

            return new SheetExecutionContext
            {
                Binding = binding,
                Definition = definition,
                Mappings = mappings,
            };
        }

        private WorksheetRuntimeColumn[] LoadRuntimeColumns(
            string sheetName,
            SheetBinding binding,
            FieldMappingTableDefinition definition,
            IReadOnlyList<SheetFieldMappingRow> mappings)
        {
            return headerMatcher.Match(sheetName, binding, definition, mappings, gridAdapter);
        }

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

                result.Add(new WorksheetRuntimeColumn
                {
                    ColumnIndex = columnIndex++,
                    ApiFieldKey = apiFieldKey,
                    HeaderType = NormalizeHeaderType(headerType),
                    DisplayText = isActivityProperty ? childText : singleText,
                    ParentDisplayText = isActivityProperty && binding.HeaderRowCount > 1 ? parentText : string.Empty,
                    ChildDisplayText = isActivityProperty ? childText : string.Empty,
                    IsIdColumn = valueAccessor.GetBoolean(definition, mapping, FieldMappingSemanticRole.IsIdColumn),
                });
            }

            return result.ToArray();
        }

        private WorksheetSchema BuildSchema(SheetBinding binding, IReadOnlyList<WorksheetRuntimeColumn> runtimeColumns)
        {
            var columns = (runtimeColumns ?? Array.Empty<WorksheetRuntimeColumn>())
                .Where(column => column != null)
                .OrderBy(column => column.ColumnIndex)
                .Select(column => new WorksheetColumnBinding
                {
                    ColumnIndex = column.ColumnIndex,
                    ApiFieldKey = column.ApiFieldKey,
                    ColumnKind = IsActivityProperty(column.HeaderType)
                        ? WorksheetColumnKind.ActivityProperty
                        : WorksheetColumnKind.Single,
                    ParentHeaderText = IsActivityProperty(column.HeaderType)
                        ? column.ParentDisplayText
                        : column.DisplayText,
                    ChildHeaderText = IsActivityProperty(column.HeaderType)
                        ? column.ChildDisplayText
                        : column.DisplayText,
                    IsIdColumn = column.IsIdColumn,
                })
                .ToArray();

            return new WorksheetSchema
            {
                SystemKey = binding?.SystemKey ?? string.Empty,
                ProjectId = binding?.ProjectId ?? string.Empty,
                Columns = columns,
            };
        }

        private bool HasAnyHeaderText(string sheetName, SheetBinding binding)
        {
            var lastUsedColumn = gridAdapter.GetLastUsedColumn(sheetName);
            if (lastUsedColumn <= 0)
            {
                return false;
            }

            var startRow = binding.HeaderStartRow;
            var rowCount = binding.HeaderRowCount;

            for (var row = startRow; row < startRow + rowCount; row++)
            {
                for (var column = 1; column <= lastUsedColumn; column++)
                {
                    if (!string.IsNullOrWhiteSpace(gridAdapter.GetCellText(sheetName, row, column)))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private void WriteFullWorksheet(WorksheetDownloadPlan plan)
        {
            var binding = plan.Binding;
            var columns = plan.RuntimeColumns ?? Array.Empty<WorksheetRuntimeColumn>();
            var clearEndRow = Math.Max(gridAdapter.GetLastUsedRow(plan.SheetName), binding.DataStartRow + (plan.Rows?.Count ?? 0) + 10);

            ClearManagedArea(plan.SheetName, binding, columns, plan.UsesExistingLayout, clearEndRow);

            if (!plan.UsesExistingLayout)
            {
                var headerPlan = layoutService.BuildHeaderPlan(binding, columns);
                foreach (var headerCell in headerPlan)
                {
                    gridAdapter.SetCellText(plan.SheetName, headerCell.Row, headerCell.Column, headerCell.Text);
                    gridAdapter.MergeCells(plan.SheetName, headerCell.Row, headerCell.Column, headerCell.RowSpan, headerCell.ColumnSpan);
                }
            }

            for (var rowIndex = 0; rowIndex < (plan.Rows?.Count ?? 0); rowIndex++)
            {
                var row = plan.Rows[rowIndex];
                var targetRow = binding.DataStartRow + rowIndex;
                foreach (var column in columns)
                {
                    gridAdapter.SetCellText(plan.SheetName, targetRow, column.ColumnIndex, GetRowValue(row, column.ApiFieldKey));
                }
            }
        }

        private void ClearManagedArea(
            string sheetName,
            SheetBinding binding,
            IReadOnlyList<WorksheetRuntimeColumn> columns,
            bool usesExistingLayout,
            int clearEndRow)
        {
            var runtimeColumns = (columns ?? Array.Empty<WorksheetRuntimeColumn>())
                .Where(column => column != null)
                .ToArray();
            if (runtimeColumns.Length == 0)
            {
                return;
            }

            var headerEndRow = binding.HeaderStartRow + Math.Max(binding.HeaderRowCount, 1) - 1;
            if (!usesExistingLayout)
            {
                var lastColumn = runtimeColumns.Max(column => column.ColumnIndex);
                gridAdapter.ClearRange(sheetName, binding.HeaderStartRow, headerEndRow, 1, lastColumn);
                gridAdapter.ClearRange(sheetName, binding.DataStartRow, clearEndRow, 1, lastColumn);
                return;
            }

            foreach (var columnIndex in runtimeColumns.Select(column => column.ColumnIndex).Distinct().OrderBy(index => index))
            {
                gridAdapter.ClearRange(sheetName, binding.DataStartRow, clearEndRow, columnIndex, columnIndex);
            }
        }

        private void WritePartialCells(WorksheetDownloadPlan plan)
        {
            var columnsByIndex = (plan.Schema?.Columns ?? Array.Empty<WorksheetColumnBinding>())
                .ToDictionary(column => column.ColumnIndex, column => column);
            var rowsById = (plan.Rows ?? Array.Empty<IDictionary<string, object>>())
                .Where(row => !string.IsNullOrWhiteSpace(GetRowId(plan.Schema, row)))
                .ToDictionary(row => GetRowId(plan.Schema, row), row => row, StringComparer.Ordinal);

            foreach (var targetCell in plan.Selection?.TargetCells ?? Array.Empty<SelectedVisibleCell>())
            {
                if (!columnsByIndex.TryGetValue(targetCell.Column, out var column))
                {
                    continue;
                }

                var rowId = GetRowId(plan.SheetName, plan.Schema, targetCell.Row);
                if (string.IsNullOrWhiteSpace(rowId) || !rowsById.TryGetValue(rowId, out var row))
                {
                    continue;
                }

                gridAdapter.SetCellText(plan.SheetName, targetCell.Row, targetCell.Column, GetRowValue(row, column.ApiFieldKey));
            }
        }

        private ResolvedSelection ResolveCurrentSelection(string sheetName, WorksheetSchema schema)
        {
            var visibleCells = selectionReader.ReadVisibleSelection() ?? Array.Empty<SelectedVisibleCell>();
            return selectionResolver.Resolve(schema, visibleCells, row => GetRowId(sheetName, schema, row));
        }

        private CellChange[] ReadAllCurrentCells(string sheetName, SheetBinding binding, WorksheetSchema schema)
        {
            var idColumn = GetIdColumn(schema);
            if (idColumn == null)
            {
                return Array.Empty<CellChange>();
            }

            var result = new List<CellChange>();
            var lastUsedRow = gridAdapter.GetLastUsedRow(sheetName);

            for (var row = binding.DataStartRow; row <= lastUsedRow; row++)
            {
                var rowId = gridAdapter.GetCellText(sheetName, row, idColumn.ColumnIndex);
                if (string.IsNullOrWhiteSpace(rowId))
                {
                    continue;
                }

                foreach (var column in schema.Columns.Where(item => !item.IsIdColumn))
                {
                    result.Add(new CellChange
                    {
                        SheetName = sheetName,
                        RowId = rowId,
                        ApiFieldKey = column.ApiFieldKey,
                        OldValue = string.Empty,
                        NewValue = gridAdapter.GetCellText(sheetName, row, column.ColumnIndex),
                    });
                }
            }

            return result.ToArray();
        }

        private CellChange[] ReadSelectionChanges(string sheetName, WorksheetSchema schema, ResolvedSelection selection)
        {
            var columnsByIndex = (schema?.Columns ?? Array.Empty<WorksheetColumnBinding>())
                .ToDictionary(column => column.ColumnIndex, column => column);
            var result = new List<CellChange>();

            foreach (var targetCell in selection?.TargetCells ?? Array.Empty<SelectedVisibleCell>())
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

                result.Add(new CellChange
                {
                    SheetName = sheetName,
                    RowId = rowId,
                    ApiFieldKey = column.ApiFieldKey,
                    OldValue = string.Empty,
                    NewValue = gridAdapter.GetCellText(sheetName, targetCell.Row, targetCell.Column),
                });
            }

            return result.ToArray();
        }

        private static IReadOnlyList<string> GetRequestedFieldKeys(IReadOnlyList<WorksheetRuntimeColumn> columns)
        {
            return (columns ?? Array.Empty<WorksheetRuntimeColumn>())
                .Where(column => column != null && !string.IsNullOrWhiteSpace(column.ApiFieldKey))
                .Select(column => column.ApiFieldKey)
                .Distinct(StringComparer.Ordinal)
                .ToArray();
        }

        private static void EnsureIdColumn(IReadOnlyList<WorksheetRuntimeColumn> columns)
        {
            if (!(columns ?? Array.Empty<WorksheetRuntimeColumn>()).Any(column => column != null && column.IsIdColumn))
            {
                throw new InvalidOperationException("SheetFieldMappings 缺少 ID 列定义，无法继续。");
            }
        }

        private bool HasUsableMappings(FieldMappingTableDefinition definition, IReadOnlyList<SheetFieldMappingRow> mappings)
        {
            var rows = mappings ?? Array.Empty<SheetFieldMappingRow>();
            if (rows.Count == 0)
            {
                return false;
            }

            var hasApiFieldKey = rows.Any(row =>
                row != null &&
                !string.IsNullOrWhiteSpace(valueAccessor.GetValue(definition, row, FieldMappingSemanticRole.ApiFieldKey)));
            if (!hasApiFieldKey)
            {
                return false;
            }

            return rows.Any(row =>
                row != null &&
                valueAccessor.GetBoolean(definition, row, FieldMappingSemanticRole.IsIdColumn) &&
                !string.IsNullOrWhiteSpace(valueAccessor.GetValue(definition, row, FieldMappingSemanticRole.ApiFieldKey)));
        }

        private static ProjectOption CreateProjectOption(SheetBinding binding)
        {
            return new ProjectOption
            {
                SystemKey = binding?.SystemKey ?? string.Empty,
                ProjectId = binding?.ProjectId ?? string.Empty,
                DisplayName = binding?.ProjectName ?? string.Empty,
            };
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
            return idColumn == null ? string.Empty : gridAdapter.GetCellText(sheetName, row, idColumn.ColumnIndex);
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

        private static InvalidOperationException CreateInitializationRequiredException()
        {
            return new InvalidOperationException("当前 sheet 未初始化，请先执行初始化当前表。");
        }

        private static InvalidOperationException CreateHeaderMatchException()
        {
            return new InvalidOperationException("当前表头无法与映射表匹配，请先修正 _Settings。");
        }

        private static void ValidateBinding(SheetBinding binding)
        {
            if (binding == null)
            {
                throw new InvalidOperationException("SheetBindings 缺少必要配置。");
            }

            if (binding.HeaderStartRow <= 0)
            {
                throw new InvalidOperationException("SheetBindings.HeaderStartRow 必须大于 0。");
            }

            if (binding.HeaderRowCount <= 0)
            {
                throw new InvalidOperationException("SheetBindings.HeaderRowCount 必须大于 0。");
            }

            if (binding.DataStartRow <= 0)
            {
                throw new InvalidOperationException("SheetBindings.DataStartRow 必须大于 0。");
            }
        }

        private static bool IsActivityProperty(string headerType)
        {
            return string.Equals(headerType, "activityProperty", StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeHeaderType(string headerType)
        {
            return string.IsNullOrWhiteSpace(headerType) ? "single" : headerType;
        }
    }

    internal sealed class SheetExecutionContext
    {
        public SheetBinding Binding { get; set; }
        public FieldMappingTableDefinition Definition { get; set; }
        public SheetFieldMappingRow[] Mappings { get; set; } = Array.Empty<SheetFieldMappingRow>();
        public WorksheetRuntimeColumn[] RuntimeColumns { get; set; } = Array.Empty<WorksheetRuntimeColumn>();
        public WorksheetSchema Schema { get; set; }
        public bool UsesExistingLayout { get; set; }
    }
}
