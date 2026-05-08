using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OfficeAgent.Core.Diagnostics;
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
        public WorksheetUploadLogCandidate[] LogCandidates { get; set; } = Array.Empty<WorksheetUploadLogCandidate>();
        public WorksheetCellAddress[] UploadedCells { get; set; } = Array.Empty<WorksheetCellAddress>();
    }

    internal sealed class WorksheetSyncExecutionService
    {
        private readonly WorksheetSyncService worksheetSyncService;
        private readonly IWorksheetSelectionReader selectionReader;
        private readonly IWorksheetGridAdapter gridAdapter;
        private readonly WorksheetSelectionResolver selectionResolver;
        private readonly WorksheetSchemaLayoutService layoutService;
        private readonly WorksheetColumnSegmentBuilder segmentBuilder;
        private readonly SyncOperationPreviewFactory previewFactory;
        private readonly WorksheetHeaderMatcher headerMatcher;
        private readonly WorksheetHeaderScanner headerScanner;
        private readonly AiColumnMappingService aiColumnMappingService;
        private readonly IAiColumnMappingClient aiColumnMappingClient;
        private readonly FieldMappingValueAccessor valueAccessor;
        private readonly ExcelUploadValueNormalizer uploadValueNormalizer;
        private readonly IWorksheetChangeLogStore changeLogStore;
        private readonly WorksheetPendingEditTracker pendingEditTracker;

        public WorksheetSyncExecutionService(
            WorksheetSyncService worksheetSyncService,
            IWorksheetMetadataStore metadataStore,
            IWorksheetSelectionReader selectionReader,
            IWorksheetGridAdapter gridAdapter,
            SyncOperationPreviewFactory previewFactory)
            : this(
                worksheetSyncService,
                metadataStore,
                selectionReader,
                gridAdapter,
                previewFactory,
                null,
                null)
        {
        }

        public WorksheetSyncExecutionService(
            WorksheetSyncService worksheetSyncService,
            IWorksheetMetadataStore metadataStore,
            IWorksheetSelectionReader selectionReader,
            IWorksheetGridAdapter gridAdapter,
            SyncOperationPreviewFactory previewFactory,
            IWorksheetChangeLogStore changeLogStore,
            WorksheetPendingEditTracker pendingEditTracker)
        {
            this.worksheetSyncService = worksheetSyncService ?? throw new ArgumentNullException(nameof(worksheetSyncService));
            _ = metadataStore ?? throw new ArgumentNullException(nameof(metadataStore));
            this.selectionReader = selectionReader ?? throw new ArgumentNullException(nameof(selectionReader));
            this.gridAdapter = gridAdapter ?? throw new ArgumentNullException(nameof(gridAdapter));
            this.previewFactory = previewFactory ?? throw new ArgumentNullException(nameof(previewFactory));
            this.changeLogStore = changeLogStore;
            this.pendingEditTracker = pendingEditTracker;
            selectionResolver = new WorksheetSelectionResolver();
            layoutService = new WorksheetSchemaLayoutService();
            segmentBuilder = new WorksheetColumnSegmentBuilder();
            valueAccessor = new FieldMappingValueAccessor();
            headerMatcher = new WorksheetHeaderMatcher(valueAccessor);
            headerScanner = new WorksheetHeaderScanner();
            aiColumnMappingService = new AiColumnMappingService();
            uploadValueNormalizer = new ExcelUploadValueNormalizer();
        }

        public WorksheetSyncExecutionService(
            WorksheetSyncService worksheetSyncService,
            IWorksheetMetadataStore metadataStore,
            IWorksheetSelectionReader selectionReader,
            IWorksheetGridAdapter gridAdapter,
            SyncOperationPreviewFactory previewFactory,
            IAiColumnMappingClient aiColumnMappingClient)
            : this(
                worksheetSyncService,
                metadataStore,
                selectionReader,
                gridAdapter,
                previewFactory)
        {
            this.aiColumnMappingClient = aiColumnMappingClient ?? throw new ArgumentNullException(nameof(aiColumnMappingClient));
        }

        public WorksheetSyncExecutionService(
            WorksheetSyncService worksheetSyncService,
            IWorksheetMetadataStore metadataStore,
            IWorksheetSelectionReader selectionReader,
            IWorksheetGridAdapter gridAdapter,
            SyncOperationPreviewFactory previewFactory,
            IWorksheetChangeLogStore changeLogStore,
            WorksheetPendingEditTracker pendingEditTracker,
            IAiColumnMappingClient aiColumnMappingClient)
            : this(
                worksheetSyncService,
                metadataStore,
                selectionReader,
                gridAdapter,
                previewFactory,
                changeLogStore,
                pendingEditTracker)
        {
            this.aiColumnMappingClient = aiColumnMappingClient ?? throw new ArgumentNullException(nameof(aiColumnMappingClient));
        }

        public void InitializeCurrentSheet(string sheetName, ProjectOption project)
        {
            worksheetSyncService.InitializeSheet(sheetName, project);
        }

        public AiColumnMappingPreview PrepareAiColumnMappingPreview(string sheetName)
        {
            if (aiColumnMappingClient == null)
            {
                throw new InvalidOperationException("AI column mapping is not configured.");
            }

            var context = PrepareAiColumnMappingContext(sheetName);
            if (context.Request.ActualHeaders == null || context.Request.ActualHeaders.Length == 0)
            {
                return new AiColumnMappingPreview();
            }

            var response = aiColumnMappingClient.Map(context.Request);
            return aiColumnMappingService.CreatePreview(context.Request, response, context.HeaderRowCount);
        }

        public async Task<AiColumnMappingPreview> PrepareAiColumnMappingPreviewAsync(
            string sheetName,
            CancellationToken cancellationToken)
        {
            if (aiColumnMappingClient == null)
            {
                throw new InvalidOperationException("AI column mapping is not configured.");
            }

            cancellationToken.ThrowIfCancellationRequested();
            var context = PrepareAiColumnMappingContext(sheetName);
            if (context.Request.ActualHeaders == null || context.Request.ActualHeaders.Length == 0)
            {
                return new AiColumnMappingPreview();
            }

            var response = await aiColumnMappingClient.MapAsync(context.Request, cancellationToken);
            return aiColumnMappingService.CreatePreview(context.Request, response, context.HeaderRowCount);
        }

        private AiColumnMappingExecutionContext PrepareAiColumnMappingContext(string sheetName)
        {
            var context = LoadSheetContext(sheetName);
            var actualHeaders = headerScanner.Scan(sheetName, context.Binding, gridAdapter);
            if (actualHeaders.Length == 0)
            {
                throw new InvalidOperationException("No header text was found in the configured header area. Check HeaderStartRow and HeaderRowCount.");
            }

            return new AiColumnMappingExecutionContext
            {
                Request = aiColumnMappingService.BuildRequest(sheetName, context.Definition, context.Mappings, actualHeaders),
                HeaderRowCount = context.Binding.HeaderRowCount,
            };
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
            worksheetSyncService.SaveFieldMappings(sheetName, context.Definition, result.Rows);
            return result;
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
            var rowIdAccessor = CreateCachedRowIdAccessor(sheetName, context.Schema);
            var selection = ResolveCurrentSelection(context.Schema, rowIdAccessor);
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
            WorksheetUploadReadResult readResult;
            using (gridAdapter.BeginBulkOperation())
            {
                readResult = ReadAllCurrentCells(sheetName, context.Binding, context.Schema);
            }

            var preview = BuildUploadPreview("全量上传", context.Binding.SystemKey, context.Binding.ProjectId, readResult.Changes);
            return new WorksheetUploadPlan
            {
                OperationName = "全量上传",
                SheetName = sheetName,
                ProjectId = context.Binding.ProjectId,
                SystemKey = context.Binding.SystemKey,
                Preview = preview,
                LogCandidates = SelectIncludedLogCandidates(readResult.LogCandidates, preview.Changes),
                UploadedCells = SelectIncludedUploadedCells(readResult, preview.Changes),
            };
        }

        public WorksheetUploadPlan PreparePartialUpload(string sheetName)
        {
            var context = ResolveMatchedSheetContext(sheetName);
            var rowIdAccessor = CreateCachedRowIdAccessor(sheetName, context.Schema);
            var selection = ResolveCurrentSelection(context.Schema, rowIdAccessor);
            var readResult = ReadSelectionChanges(sheetName, context.Schema, selection, rowIdAccessor);

            var preview = BuildUploadPreview("部分上传", context.Binding.SystemKey, context.Binding.ProjectId, readResult.Changes);
            return new WorksheetUploadPlan
            {
                OperationName = "部分上传",
                SheetName = sheetName,
                ProjectId = context.Binding.ProjectId,
                SystemKey = context.Binding.SystemKey,
                Preview = preview,
                LogCandidates = SelectIncludedLogCandidates(readResult.LogCandidates, preview.Changes),
                UploadedCells = SelectIncludedUploadedCells(readResult, preview.Changes),
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
            AppendChangeLogEntries(BuildUploadLogEntries(plan.LogCandidates));
            ClearPendingUploadedCells(plan);
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
                throw CreateInitializationRequiredException();
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
                var isSingleHeader = string.IsNullOrWhiteSpace(headerType) ||
                                     string.Equals(headerType, "single", StringComparison.OrdinalIgnoreCase);
                var isGroupedSingle = isSingleHeader && !isActivityProperty && !string.IsNullOrWhiteSpace(childText);

                result.Add(new WorksheetRuntimeColumn
                {
                    ColumnIndex = columnIndex++,
                    ApiFieldKey = apiFieldKey,
                    HeaderType = NormalizeHeaderType(headerType),
                    DisplayText = isActivityProperty
                        ? childText
                        : (isGroupedSingle ? childText : singleText),
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
            var logEntries = BuildFullDownloadLogEntries(plan);

            using (gridAdapter.BeginBulkOperation())
            {
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

                WriteFullDataRows(plan.SheetName, binding.DataStartRow, columns, plan.Rows);
            }

            AppendChangeLogEntries(logEntries);
        }

        private void WriteFullDataRows(
            string sheetName,
            int startRow,
            IReadOnlyList<WorksheetRuntimeColumn> columns,
            IReadOnlyList<IDictionary<string, object>> rows)
        {
            var sourceRows = rows ?? Array.Empty<IDictionary<string, object>>();
            if (sourceRows.Count == 0)
            {
                return;
            }

            foreach (var segment in segmentBuilder.Build(columns))
            {
                if (segment?.Columns == null || segment.Columns.Length == 0)
                {
                    continue;
                }

                var values = new object[sourceRows.Count, segment.Columns.Length];
                for (var rowIndex = 0; rowIndex < sourceRows.Count; rowIndex++)
                {
                    var row = sourceRows[rowIndex];
                    for (var columnOffset = 0; columnOffset < segment.Columns.Length; columnOffset++)
                    {
                        values[rowIndex, columnOffset] = GetRowValue(row, segment.Columns[columnOffset].ApiFieldKey);
                    }
                }

                gridAdapter.WriteRangeValues(sheetName, startRow, segment.StartColumn, values);
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
            var writeBatches = BuildPartialWriteBatches(plan, out var writeCells);
            if (writeBatches.Count == 0)
            {
                return;
            }

            using (gridAdapter.BeginBulkOperation())
            {
                foreach (var batch in writeBatches)
                {
                    gridAdapter.WriteRangeValues(plan.SheetName, batch.StartRow, batch.StartColumn, batch.Values);
                }
            }

            AppendChangeLogEntries(BuildDownloadLogEntries(writeCells));
        }

        private List<PartialWorksheetWriteBatch> BuildPartialWriteBatches(
            WorksheetDownloadPlan plan,
            out List<PartialWorksheetWriteCell> writeCells)
        {
            var columnsByIndex = (plan.Schema?.Columns ?? Array.Empty<WorksheetColumnBinding>())
                .ToDictionary(column => column.ColumnIndex, column => column);
            var rowsById = (plan.Rows ?? Array.Empty<IDictionary<string, object>>())
                .Where(row => !string.IsNullOrWhiteSpace(GetRowId(plan.Schema, row)))
                .ToDictionary(row => GetRowId(plan.Schema, row), row => row, StringComparer.Ordinal);
            var rowIdAccessor = CreateCachedRowIdAccessor(plan.SheetName, plan.Schema);
            writeCells = new List<PartialWorksheetWriteCell>();

            foreach (var targetCell in plan.Selection?.TargetCells ?? Array.Empty<SelectedVisibleCell>())
            {
                if (!columnsByIndex.TryGetValue(targetCell.Column, out var column))
                {
                    continue;
                }

                var rowId = rowIdAccessor(targetCell.Row);
                if (string.IsNullOrWhiteSpace(rowId) || !rowsById.TryGetValue(rowId, out var row))
                {
                    continue;
                }

                var value = GetRowValue(row, column.ApiFieldKey);
                writeCells.Add(new PartialWorksheetWriteCell
                {
                    Row = targetCell.Row,
                    Column = targetCell.Column,
                    RowId = rowId,
                    HeaderText = BuildHeaderText(column),
                    IsIdColumn = column.IsIdColumn,
                    OldValue = gridAdapter.GetCellText(plan.SheetName, targetCell.Row, targetCell.Column),
                    Value = value,
                });
            }

            if (writeCells.Count == 0)
            {
                return new List<PartialWorksheetWriteBatch>();
            }

            var rowSegments = BuildPartialRowSegments(writeCells);
            var batches = new List<PartialWorksheetWriteBatch>();
            PartialWorksheetWriteBatch currentBatch = null;

            foreach (var segment in rowSegments)
            {
                if (currentBatch != null &&
                    segment.Row == currentBatch.EndRow + 1 &&
                    segment.StartColumn == currentBatch.StartColumn &&
                    segment.EndColumn == currentBatch.EndColumn)
                {
                    currentBatch.EndRow = segment.Row;
                    currentBatch.RowValues.Add(segment.Values);
                    continue;
                }

                if (currentBatch != null)
                {
                    batches.Add(FinalizePartialWriteBatch(currentBatch));
                }

                currentBatch = new PartialWorksheetWriteBatch
                {
                    StartRow = segment.Row,
                    EndRow = segment.Row,
                    StartColumn = segment.StartColumn,
                    EndColumn = segment.EndColumn,
                    RowValues = new List<object[]> { segment.Values },
                };
            }

            if (currentBatch != null)
            {
                batches.Add(FinalizePartialWriteBatch(currentBatch));
            }

            return batches;
        }

        private WorksheetChangeLogEntry[] BuildFullDownloadLogEntries(WorksheetDownloadPlan plan)
        {
            var binding = plan?.Binding;
            var columns = (plan?.RuntimeColumns ?? Array.Empty<WorksheetRuntimeColumn>())
                .Where(column => column != null && !column.IsIdColumn)
                .ToArray();
            var rows = plan?.Rows ?? Array.Empty<IDictionary<string, object>>();
            if (binding == null || columns.Length == 0 || rows.Count == 0)
            {
                return Array.Empty<WorksheetChangeLogEntry>();
            }

            var entries = new List<WorksheetChangeLogEntry>();
            for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                var row = rows[rowIndex];
                var rowId = GetRowId(plan.Schema, row);
                if (string.IsNullOrWhiteSpace(rowId))
                {
                    continue;
                }

                foreach (var column in columns)
                {
                    var oldValue = gridAdapter.GetCellText(plan.SheetName, binding.DataStartRow + rowIndex, column.ColumnIndex);
                    var newValue = GetRowValue(row, column.ApiFieldKey);
                    if (string.Equals(oldValue, newValue, StringComparison.Ordinal))
                    {
                        continue;
                    }

                    entries.Add(new WorksheetChangeLogEntry
                    {
                        Key = rowId,
                        HeaderText = BuildHeaderText(column),
                        ChangeMode = "下载",
                        NewValue = newValue,
                        OldValue = oldValue,
                    });
                }
            }

            return entries.ToArray();
        }

        private static WorksheetChangeLogEntry[] BuildDownloadLogEntries(IEnumerable<PartialWorksheetWriteCell> writeCells)
        {
            return (writeCells ?? Enumerable.Empty<PartialWorksheetWriteCell>())
                .Where(cell => cell != null && !cell.IsIdColumn)
                .Where(cell => !string.IsNullOrWhiteSpace(cell.RowId))
                .Select(cell => new
                {
                    Cell = cell,
                    NewValue = Convert.ToString(cell.Value) ?? string.Empty,
                })
                .Where(item => !string.Equals(item.Cell.OldValue ?? string.Empty, item.NewValue, StringComparison.Ordinal))
                .Select(item => new WorksheetChangeLogEntry
                {
                    Key = item.Cell.RowId,
                    HeaderText = item.Cell.HeaderText,
                    ChangeMode = "下载",
                    NewValue = item.NewValue,
                    OldValue = item.Cell.OldValue ?? string.Empty,
                })
                .ToArray();
        }

        private WorksheetUploadLogCandidate TryCreateUploadLogCandidate(
            string sheetName,
            int row,
            WorksheetColumnBinding column,
            string rowId,
            string newValue)
        {
            if (pendingEditTracker == null ||
                column == null ||
                column.IsIdColumn ||
                string.IsNullOrWhiteSpace(rowId))
            {
                return null;
            }

            if (!pendingEditTracker.TryGetOriginalValue(sheetName, row, column.ColumnIndex, out var oldValue))
            {
                return null;
            }

            if (string.Equals(oldValue ?? string.Empty, newValue ?? string.Empty, StringComparison.Ordinal))
            {
                return null;
            }

            return new WorksheetUploadLogCandidate
            {
                Row = row,
                Column = column.ColumnIndex,
                RowId = rowId,
                ApiFieldKey = column.ApiFieldKey,
                HeaderText = BuildHeaderText(column),
                NewValue = newValue ?? string.Empty,
                OldValue = oldValue ?? string.Empty,
            };
        }

        private static WorksheetChangeLogEntry[] BuildUploadLogEntries(
            IReadOnlyList<WorksheetUploadLogCandidate> candidates)
        {
            return (candidates ?? Array.Empty<WorksheetUploadLogCandidate>())
                .Where(candidate => candidate != null)
                .Select(candidate => new WorksheetChangeLogEntry
                {
                    Key = candidate.RowId,
                    HeaderText = candidate.HeaderText,
                    ChangeMode = "上传",
                    NewValue = candidate.NewValue,
                    OldValue = candidate.OldValue,
                })
                .ToArray();
        }

        private void ClearPendingUploadedCells(WorksheetUploadPlan plan)
        {
            if (pendingEditTracker == null || plan == null)
            {
                return;
            }

            pendingEditTracker.Clear(plan.SheetName, plan.UploadedCells);
        }

        private void AppendChangeLogEntries(IReadOnlyList<WorksheetChangeLogEntry> entries)
        {
            if (changeLogStore == null || entries == null || entries.Count == 0)
            {
                return;
            }

            try
            {
                changeLogStore.Append(entries);
            }
            catch (Exception exception)
            {
                OfficeAgentLog.Error(
                    "ribbon-sync",
                    "change-log.append.failed",
                    "Failed to append workbook change log entries.",
                    exception);
            }
        }

        private static List<PartialWorksheetWriteRowSegment> BuildPartialRowSegments(IEnumerable<PartialWorksheetWriteCell> writeCells)
        {
            var normalizedCells = (writeCells ?? Enumerable.Empty<PartialWorksheetWriteCell>())
                .Where(cell => cell != null)
                .GroupBy(cell => new { cell.Row, cell.Column })
                .Select(group => group.Last())
                .OrderBy(cell => cell.Row)
                .ThenBy(cell => cell.Column);
            var segments = new List<PartialWorksheetWriteRowSegment>();

            foreach (var rowGroup in normalizedCells.GroupBy(cell => cell.Row))
            {
                PartialWorksheetWriteRowSegment currentSegment = null;
                foreach (var cell in rowGroup.OrderBy(item => item.Column))
                {
                    if (currentSegment != null && cell.Column == currentSegment.EndColumn + 1)
                    {
                        currentSegment.EndColumn = cell.Column;
                        currentSegment.Cells.Add(cell.Value);
                        continue;
                    }

                    if (currentSegment != null)
                    {
                        segments.Add(new PartialWorksheetWriteRowSegment
                        {
                            Row = currentSegment.Row,
                            StartColumn = currentSegment.StartColumn,
                            EndColumn = currentSegment.EndColumn,
                            Values = currentSegment.Cells.ToArray(),
                        });
                    }

                    currentSegment = new PartialWorksheetWriteRowSegment
                    {
                        Row = cell.Row,
                        StartColumn = cell.Column,
                        EndColumn = cell.Column,
                        Cells = new List<object> { cell.Value },
                    };
                }

                if (currentSegment != null)
                {
                    segments.Add(new PartialWorksheetWriteRowSegment
                    {
                        Row = currentSegment.Row,
                        StartColumn = currentSegment.StartColumn,
                        EndColumn = currentSegment.EndColumn,
                        Values = currentSegment.Cells.ToArray(),
                    });
                }
            }

            return segments;
        }

        private static PartialWorksheetWriteBatch FinalizePartialWriteBatch(PartialWorksheetWriteBatch batch)
        {
            if (batch == null)
            {
                return null;
            }

            var rowValues = batch.RowValues ?? new List<object[]>();
            var columnCount = rowValues.Count == 0 ? 0 : rowValues[0]?.Length ?? 0;
            var values = new object[rowValues.Count, columnCount];

            for (var rowIndex = 0; rowIndex < rowValues.Count; rowIndex++)
            {
                var sourceRow = rowValues[rowIndex] ?? Array.Empty<object>();
                for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
                {
                    values[rowIndex, columnIndex] = columnIndex < sourceRow.Length
                        ? sourceRow[columnIndex]
                        : string.Empty;
                }
            }

            batch.Values = values;
            return batch;
        }

        private ResolvedSelection ResolveCurrentSelection(WorksheetSchema schema, Func<int, string> rowIdAccessor)
        {
            var visibleCells = selectionReader.ReadVisibleSelection() ?? Array.Empty<SelectedVisibleCell>();
            return selectionResolver.Resolve(schema, visibleCells, rowIdAccessor);
        }

        private WorksheetUploadReadResult ReadAllCurrentCells(string sheetName, SheetBinding binding, WorksheetSchema schema)
        {
            var idColumn = GetIdColumn(schema);
            if (idColumn == null)
            {
                return WorksheetUploadReadResult.Empty;
            }

            var result = new List<CellChange>();
            var logCandidates = new List<WorksheetUploadLogCandidate>();
            var uploadedCells = new List<WorksheetCellAddress>();
            var lastUsedRow = gridAdapter.GetLastUsedRow(sheetName);
            if (lastUsedRow < binding.DataStartRow)
            {
                return WorksheetUploadReadResult.Empty;
            }

            var columns = (schema.Columns ?? Array.Empty<WorksheetColumnBinding>())
                .Where(column => column != null)
                .OrderBy(column => column.ColumnIndex)
                .ToArray();
            if (columns.Length == 0)
            {
                return WorksheetUploadReadResult.Empty;
            }

            var segments = ReadUploadSegments(sheetName, binding.DataStartRow, lastUsedRow, columns);
            var nonIdColumns = columns.Where(column => !column.IsIdColumn).ToArray();

            for (var row = binding.DataStartRow; row <= lastUsedRow; row++)
            {
                var rowOffset = row - binding.DataStartRow;
                var rowId = ReadUploadCellValue(sheetName, row, idColumn.ColumnIndex, rowOffset, segments);
                if (string.IsNullOrWhiteSpace(rowId))
                {
                    continue;
                }

                foreach (var column in nonIdColumns)
                {
                    var newValue = ReadUploadCellValue(sheetName, row, column.ColumnIndex, rowOffset, segments);
                    result.Add(new CellChange
                    {
                        SheetName = sheetName,
                        RowId = rowId,
                        ApiFieldKey = column.ApiFieldKey,
                        OldValue = string.Empty,
                        NewValue = newValue,
                    });
                    uploadedCells.Add(new WorksheetCellAddress { Row = row, Column = column.ColumnIndex });

                    var logCandidate = TryCreateUploadLogCandidate(sheetName, row, column, rowId, newValue);
                    if (logCandidate != null)
                    {
                        logCandidates.Add(logCandidate);
                    }
                }
            }

            return new WorksheetUploadReadResult
            {
                Changes = result.ToArray(),
                LogCandidates = logCandidates.ToArray(),
                UploadedCells = uploadedCells.ToArray(),
            };
        }

        private WorksheetUploadReadSegment[] ReadUploadSegments(
            string sheetName,
            int startRow,
            int endRow,
            IReadOnlyList<WorksheetColumnBinding> columns)
        {
            var segmentColumns = (columns ?? Array.Empty<WorksheetColumnBinding>())
                .Where(column => column != null)
                .Select(column => new WorksheetRuntimeColumn
                {
                    ColumnIndex = column.ColumnIndex,
                    ApiFieldKey = column.ApiFieldKey,
                    IsIdColumn = column.IsIdColumn,
                })
                .ToArray();

            return segmentBuilder.Build(segmentColumns)
                .Select(segment => new WorksheetUploadReadSegment
                {
                    StartColumn = segment.StartColumn,
                    EndColumn = segment.EndColumn,
                    Values = gridAdapter.ReadRangeValues(sheetName, startRow, endRow, segment.StartColumn, segment.EndColumn),
                    NumberFormats = gridAdapter.ReadRangeNumberFormats(sheetName, startRow, endRow, segment.StartColumn, segment.EndColumn),
                })
                .ToArray();
        }

        private WorksheetUploadReadResult ReadSelectionChanges(
            string sheetName,
            WorksheetSchema schema,
            ResolvedSelection selection,
            Func<int, string> rowIdAccessor)
        {
            var columnsByIndex = (schema?.Columns ?? Array.Empty<WorksheetColumnBinding>())
                .ToDictionary(column => column.ColumnIndex, column => column);
            var result = new List<CellChange>();
            var logCandidates = new List<WorksheetUploadLogCandidate>();
            var uploadedCells = new List<WorksheetCellAddress>();

            foreach (var targetCell in selection?.TargetCells ?? Array.Empty<SelectedVisibleCell>())
            {
                if (!columnsByIndex.TryGetValue(targetCell.Column, out var column) || column.IsIdColumn)
                {
                    continue;
                }

                var rowId = rowIdAccessor(targetCell.Row);
                if (string.IsNullOrWhiteSpace(rowId))
                {
                    continue;
                }

                var newValue = gridAdapter.GetCellText(sheetName, targetCell.Row, targetCell.Column);
                result.Add(new CellChange
                {
                    SheetName = sheetName,
                    RowId = rowId,
                    ApiFieldKey = column.ApiFieldKey,
                    OldValue = string.Empty,
                    NewValue = newValue,
                });
                uploadedCells.Add(new WorksheetCellAddress { Row = targetCell.Row, Column = targetCell.Column });

                var logCandidate = TryCreateUploadLogCandidate(sheetName, targetCell.Row, column, rowId, newValue);
                if (logCandidate != null)
                {
                    logCandidates.Add(logCandidate);
                }
            }

            return new WorksheetUploadReadResult
            {
                Changes = result.ToArray(),
                LogCandidates = logCandidates.ToArray(),
                UploadedCells = uploadedCells.ToArray(),
            };
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

        private SyncOperationPreview BuildUploadPreview(
            string operationName,
            string systemKey,
            string projectId,
            IReadOnlyList<CellChange> changes)
        {
            var filterResult = worksheetSyncService.FilterUploadChanges(systemKey, projectId, changes);
            var preview = previewFactory.CreateUploadPreview(
                operationName,
                filterResult.IncludedChanges,
                filterResult.SkippedChanges);
            preview.OperationName = operationName;
            if (preview.SkippedChanges.Length == 0)
            {
                preview.Summary = $"{operationName}将提交 {preview.Changes.Length} 个单元格。";
            }

            return preview;
        }

        private static WorksheetUploadLogCandidate[] SelectIncludedLogCandidates(
            IReadOnlyList<WorksheetUploadLogCandidate> candidates,
            IReadOnlyList<CellChange> includedChanges)
        {
            var includedCounts = CreateChangeKeyCounts(includedChanges);
            return (candidates ?? Array.Empty<WorksheetUploadLogCandidate>())
                .Where(candidate => candidate != null)
                .Where(candidate => ConsumeChangeKey(
                    includedCounts,
                    candidate.RowId,
                    candidate.ApiFieldKey,
                    candidate.NewValue))
                .ToArray();
        }

        private static WorksheetCellAddress[] SelectIncludedUploadedCells(
            WorksheetUploadReadResult readResult,
            IReadOnlyList<CellChange> includedChanges)
        {
            var includedCounts = CreateChangeKeyCounts(includedChanges);
            var sourceChanges = readResult?.Changes ?? Array.Empty<CellChange>();
            var sourceCells = readResult?.UploadedCells ?? Array.Empty<WorksheetCellAddress>();
            var result = new List<WorksheetCellAddress>();

            for (var index = 0; index < sourceChanges.Length && index < sourceCells.Length; index++)
            {
                var change = sourceChanges[index];
                if (ConsumeChangeKey(includedCounts, change?.RowId, change?.ApiFieldKey, change?.NewValue))
                {
                    result.Add(sourceCells[index]);
                }
            }

            return result.ToArray();
        }

        private static Dictionary<string, int> CreateChangeKeyCounts(IReadOnlyList<CellChange> changes)
        {
            var result = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (var change in changes ?? Array.Empty<CellChange>())
            {
                if (change == null)
                {
                    continue;
                }

                var key = CreateChangeKey(change.RowId, change.ApiFieldKey, change.NewValue);
                result.TryGetValue(key, out var count);
                result[key] = count + 1;
            }

            return result;
        }

        private static bool ConsumeChangeKey(
            IDictionary<string, int> counts,
            string rowId,
            string apiFieldKey,
            string newValue)
        {
            if (counts == null)
            {
                return false;
            }

            var key = CreateChangeKey(rowId, apiFieldKey, newValue);
            if (!counts.TryGetValue(key, out var count) || count <= 0)
            {
                return false;
            }

            if (count == 1)
            {
                counts.Remove(key);
            }
            else
            {
                counts[key] = count - 1;
            }

            return true;
        }

        private static string CreateChangeKey(string rowId, string apiFieldKey, string newValue)
        {
            return $"{rowId ?? string.Empty}\u001F{apiFieldKey ?? string.Empty}\u001F{newValue ?? string.Empty}";
        }

        private static string BuildHeaderText(WorksheetColumnBinding column)
        {
            if (column == null)
            {
                return string.Empty;
            }

            var parent = column.ParentHeaderText ?? string.Empty;
            var child = column.ChildHeaderText ?? string.Empty;
            if (column.ColumnKind == WorksheetColumnKind.ActivityProperty &&
                !string.IsNullOrWhiteSpace(parent) &&
                !string.IsNullOrWhiteSpace(child))
            {
                return $"{parent}/{child}";
            }

            if (!string.IsNullOrWhiteSpace(child))
            {
                return child;
            }

            return !string.IsNullOrWhiteSpace(parent)
                ? parent
                : column.ApiFieldKey ?? string.Empty;
        }

        private static string BuildHeaderText(WorksheetRuntimeColumn column)
        {
            if (column == null)
            {
                return string.Empty;
            }

            var parent = column.ParentDisplayText ?? string.Empty;
            var child = column.ChildDisplayText ?? string.Empty;
            if (IsActivityProperty(column.HeaderType) &&
                !string.IsNullOrWhiteSpace(parent) &&
                !string.IsNullOrWhiteSpace(child))
            {
                return $"{parent}/{child}";
            }

            if (!string.IsNullOrWhiteSpace(column.DisplayText))
            {
                return column.DisplayText;
            }

            if (!string.IsNullOrWhiteSpace(child))
            {
                return child;
            }

            return !string.IsNullOrWhiteSpace(parent)
                ? parent
                : column.ApiFieldKey ?? string.Empty;
        }

        private Func<int, string> CreateCachedRowIdAccessor(string sheetName, WorksheetSchema schema)
        {
            var idColumn = GetIdColumn(schema);
            if (idColumn == null)
            {
                return _ => string.Empty;
            }

            var cache = new Dictionary<int, string>();
            return row =>
            {
                if (cache.TryGetValue(row, out var rowId))
                {
                    return rowId;
                }

                rowId = gridAdapter.GetCellText(sheetName, row, idColumn.ColumnIndex);
                cache[row] = rowId ?? string.Empty;
                return cache[row];
            };
        }

        private string ReadUploadCellValue(
            string sheetName,
            int row,
            int column,
            int rowOffset,
            IReadOnlyList<WorksheetUploadReadSegment> segments)
        {
            var segment = (segments ?? Array.Empty<WorksheetUploadReadSegment>())
                .FirstOrDefault(item => item != null && column >= item.StartColumn && column <= item.EndColumn);
            if (segment == null)
            {
                return gridAdapter.GetCellText(sheetName, row, column);
            }

            var columnOffset = column - segment.StartColumn;
            var values = segment.Values;
            var numberFormats = segment.NumberFormats;
            var value = GetRangeValue(values, rowOffset, columnOffset);
            var numberFormat = GetRangeValue(numberFormats, rowOffset, columnOffset);
            if (uploadValueNormalizer.TryNormalize(value, numberFormat, out var normalized))
            {
                return normalized;
            }

            return gridAdapter.GetCellText(sheetName, row, column);
        }

        private static object GetRangeValue(object[,] values, int rowOffset, int columnOffset)
        {
            if (values == null ||
                rowOffset < 0 ||
                columnOffset < 0 ||
                rowOffset >= values.GetLength(0) ||
                columnOffset >= values.GetLength(1))
            {
                return null;
            }

            return values[rowOffset, columnOffset];
        }

        private static string GetRangeValue(string[,] values, int rowOffset, int columnOffset)
        {
            if (values == null ||
                rowOffset < 0 ||
                columnOffset < 0 ||
                rowOffset >= values.GetLength(0) ||
                columnOffset >= values.GetLength(1))
            {
                return string.Empty;
            }

            return values[rowOffset, columnOffset] ?? string.Empty;
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
            return new InvalidOperationException($"当前表头无法与映射表匹配，请先修正 {MetadataWorksheetNames.Current}。");
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

    internal sealed class AiColumnMappingExecutionContext
    {
        public AiColumnMappingRequest Request { get; set; }

        public int HeaderRowCount { get; set; }
    }

    internal sealed class WorksheetUploadReadSegment
    {
        public int StartColumn { get; set; }
        public int EndColumn { get; set; }
        public object[,] Values { get; set; }
        public string[,] NumberFormats { get; set; }
    }

    internal sealed class WorksheetUploadReadResult
    {
        public static readonly WorksheetUploadReadResult Empty = new WorksheetUploadReadResult();

        public CellChange[] Changes { get; set; } = Array.Empty<CellChange>();

        public WorksheetUploadLogCandidate[] LogCandidates { get; set; } = Array.Empty<WorksheetUploadLogCandidate>();

        public WorksheetCellAddress[] UploadedCells { get; set; } = Array.Empty<WorksheetCellAddress>();
    }

    internal sealed class WorksheetUploadLogCandidate
    {
        public int Row { get; set; }

        public int Column { get; set; }

        public string RowId { get; set; } = string.Empty;

        public string ApiFieldKey { get; set; } = string.Empty;

        public string HeaderText { get; set; } = string.Empty;

        public string NewValue { get; set; } = string.Empty;

        public string OldValue { get; set; } = string.Empty;
    }

    internal sealed class PartialWorksheetWriteCell
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public string RowId { get; set; } = string.Empty;
        public string HeaderText { get; set; } = string.Empty;
        public bool IsIdColumn { get; set; }
        public string OldValue { get; set; } = string.Empty;
        public object Value { get; set; }
    }

    internal sealed class PartialWorksheetWriteRowSegment
    {
        public int Row { get; set; }
        public int StartColumn { get; set; }
        public int EndColumn { get; set; }
        public object[] Values { get; set; } = Array.Empty<object>();
        public List<object> Cells { get; set; } = new List<object>();
    }

    internal sealed class PartialWorksheetWriteBatch
    {
        public int StartRow { get; set; }
        public int EndRow { get; set; }
        public int StartColumn { get; set; }
        public int EndColumn { get; set; }
        public List<object[]> RowValues { get; set; } = new List<object[]>();
        public object[,] Values { get; set; } = new object[0, 0];
    }
}
