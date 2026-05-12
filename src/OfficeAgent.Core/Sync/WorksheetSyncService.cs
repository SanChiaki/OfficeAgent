using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using OfficeAgent.Core.Analytics;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;

namespace OfficeAgent.Core.Sync
{
    public sealed class WorksheetSyncService
    {
        private readonly ISystemConnectorRegistry connectorRegistry;
        private readonly IWorksheetMetadataStore metadataStore;
        private readonly IAnalyticsService analyticsService;

        public WorksheetSyncService(
            ISystemConnectorRegistry connectorRegistry,
            IWorksheetMetadataStore metadataStore,
            IAnalyticsService analyticsService = null)
        {
            this.connectorRegistry = connectorRegistry ?? throw new ArgumentNullException(nameof(connectorRegistry));
            this.metadataStore = metadataStore ?? throw new ArgumentNullException(nameof(metadataStore));
            this.analyticsService = analyticsService ?? NoopAnalyticsService.Instance;
        }

        public WorksheetSyncService(
            ISystemConnectorRegistry connectorRegistry,
            IWorksheetMetadataStore metadataStore,
            WorksheetChangeTracker changeTracker,
            SyncOperationPreviewFactory previewFactory,
            IAnalyticsService analyticsService = null)
            : this(connectorRegistry, metadataStore, analyticsService)
        {
        }

        public IReadOnlyList<ProjectOption> GetProjects()
        {
            var stopwatch = Stopwatch.StartNew();
            var properties = BuildConnectorProperties(string.Empty, string.Empty);
            try
            {
                var projects = connectorRegistry.GetProjects() ?? Array.Empty<ProjectOption>();
                properties["projectCount"] = projects.Count;
                TrackConnectorEvent("connector.projects.completed", properties, stopwatch);
                return projects;
            }
            catch (Exception ex)
            {
                TrackConnectorEvent("connector.projects.failed", properties, stopwatch, ToAnalyticsError(ex));
                throw;
            }
        }

        public void InitializeSheet(string sheetName, ProjectOption project)
        {
            var stopwatch = Stopwatch.StartNew();
            var properties = BuildConnectorProperties(project?.SystemKey, project?.ProjectId);
            try
            {
                if (string.IsNullOrWhiteSpace(sheetName))
                {
                    throw new ArgumentException("Sheet name is required.", nameof(sheetName));
                }

                if (project == null)
                {
                    throw new ArgumentNullException(nameof(project));
                }

                var connector = GetRequiredConnector(project.SystemKey);
                var bindingSeed = connector.CreateBindingSeed(sheetName, project);
                var binding = MergeExistingLayout(bindingSeed);
                var definition = connector.GetFieldMappingDefinition(project.ProjectId);
                var seedRows = connector.BuildFieldMappingSeed(sheetName, project.ProjectId);

                metadataStore.SaveBinding(binding);
                metadataStore.SaveFieldMappings(sheetName, definition, seedRows);
                properties["fieldMappingColumnCount"] = definition?.Columns?.Length ?? 0;
                properties["fieldMappingRowCount"] = seedRows?.Count ?? 0;
                TrackConnectorEvent("connector.initialize_sheet.completed", properties, stopwatch);
            }
            catch (Exception ex)
            {
                TrackConnectorEvent("connector.initialize_sheet.failed", properties, stopwatch, ToAnalyticsError(ex));
                throw;
            }
        }

        public SheetBinding CreateBindingSeed(string sheetName, ProjectOption project)
        {
            var stopwatch = Stopwatch.StartNew();
            var properties = BuildConnectorProperties(project?.SystemKey, project?.ProjectId);
            try
            {
                if (string.IsNullOrWhiteSpace(sheetName))
                {
                    throw new ArgumentException("Sheet name is required.", nameof(sheetName));
                }

                if (project == null)
                {
                    throw new ArgumentNullException(nameof(project));
                }

                var connector = GetRequiredConnector(project.SystemKey);
                var binding = MergeExistingLayout(connector.CreateBindingSeed(sheetName, project));
                TrackConnectorEvent("connector.binding_seed.completed", properties, stopwatch);
                return binding;
            }
            catch (Exception ex)
            {
                TrackConnectorEvent("connector.binding_seed.failed", properties, stopwatch, ToAnalyticsError(ex));
                throw;
            }
        }

        public SheetBinding LoadBinding(string sheetName)
        {
            return metadataStore.LoadBinding(sheetName);
        }

        public FieldMappingTableDefinition LoadFieldMappingDefinition(string systemKey, string projectId)
        {
            var stopwatch = Stopwatch.StartNew();
            var properties = BuildConnectorProperties(systemKey, projectId);
            try
            {
                var definition = GetRequiredConnector(systemKey).GetFieldMappingDefinition(projectId);
                properties["fieldMappingColumnCount"] = definition?.Columns?.Length ?? 0;
                TrackConnectorEvent("connector.field_mapping_definition.completed", properties, stopwatch);
                return definition;
            }
            catch (Exception ex)
            {
                TrackConnectorEvent("connector.field_mapping_definition.failed", properties, stopwatch, ToAnalyticsError(ex));
                throw;
            }
        }

        public SheetFieldMappingRow[] LoadFieldMappings(string sheetName, string systemKey, string projectId)
        {
            var definition = LoadFieldMappingDefinition(systemKey, projectId);
            return metadataStore.LoadFieldMappings(sheetName, definition);
        }

        public void SaveFieldMappings(
            string sheetName,
            FieldMappingTableDefinition definition,
            IReadOnlyList<SheetFieldMappingRow> rows)
        {
            metadataStore.SaveFieldMappings(sheetName, definition, rows);
        }

        public IReadOnlyList<IDictionary<string, object>> Download(
            string systemKey,
            string projectId,
            IReadOnlyList<string> rowIds,
            IReadOnlyList<string> fieldKeys)
        {
            var stopwatch = Stopwatch.StartNew();
            var properties = BuildConnectorProperties(systemKey, projectId);
            properties["rowIdCount"] = rowIds?.Count ?? 0;
            properties["fieldKeyCount"] = fieldKeys?.Count ?? 0;
            try
            {
                var rows = GetRequiredConnector(systemKey).Find(projectId, rowIds, fieldKeys);
                properties["resultRowCount"] = rows?.Count ?? 0;
                TrackConnectorEvent("connector.find.completed", properties, stopwatch);
                return rows;
            }
            catch (Exception ex)
            {
                TrackConnectorEvent("connector.find.failed", properties, stopwatch, ToAnalyticsError(ex));
                throw;
            }
        }

        public void Upload(string systemKey, string projectId, IReadOnlyList<CellChange> changes)
        {
            var stopwatch = Stopwatch.StartNew();
            var properties = BuildConnectorProperties(systemKey, projectId);
            properties["changeCount"] = changes?.Count ?? 0;
            try
            {
                GetRequiredConnector(systemKey).BatchSave(projectId, changes);
                TrackConnectorEvent("connector.batch_save.completed", properties, stopwatch);
            }
            catch (Exception ex)
            {
                TrackConnectorEvent("connector.batch_save.failed", properties, stopwatch, ToAnalyticsError(ex));
                throw;
            }
        }

        public UploadChangeFilterResult FilterUploadChanges(
            string systemKey,
            string projectId,
            IReadOnlyList<CellChange> changes)
        {
            var stopwatch = Stopwatch.StartNew();
            var properties = BuildConnectorProperties(systemKey, projectId);
            properties["changeCount"] = changes?.Count ?? 0;
            try
            {
                var changeList = changes ?? Array.Empty<CellChange>();
                var connector = GetRequiredConnector(systemKey);
                UploadChangeFilterResult normalizedResult;
                if (connector is IUploadChangeFilter filter)
                {
                    var result = filter.FilterUploadChanges(projectId, changeList);
                    if (result != null)
                    {
                        normalizedResult = new UploadChangeFilterResult
                        {
                            IncludedChanges = result.IncludedChanges ?? Array.Empty<CellChange>(),
                            SkippedChanges = result.SkippedChanges ?? Array.Empty<SkippedCellChange>(),
                        };
                        properties["includedCount"] = normalizedResult.IncludedChanges.Length;
                        properties["skippedCount"] = normalizedResult.SkippedChanges.Length;
                        TrackConnectorEvent("connector.upload_filter.completed", properties, stopwatch);
                        return normalizedResult;
                    }
                }

                normalizedResult = new UploadChangeFilterResult
                {
                    IncludedChanges = changeList.ToArray(),
                    SkippedChanges = Array.Empty<SkippedCellChange>(),
                };
                properties["includedCount"] = normalizedResult.IncludedChanges.Length;
                properties["skippedCount"] = normalizedResult.SkippedChanges.Length;
                TrackConnectorEvent("connector.upload_filter.completed", properties, stopwatch);
                return normalizedResult;
            }
            catch (Exception ex)
            {
                TrackConnectorEvent("connector.upload_filter.failed", properties, stopwatch, ToAnalyticsError(ex));
                throw;
            }
        }

        private ISystemConnector GetRequiredConnector(string systemKey)
        {
            return connectorRegistry.GetRequiredConnector(systemKey);
        }

        private void TrackConnectorEvent(
            string eventName,
            Dictionary<string, object> properties,
            Stopwatch stopwatch,
            AnalyticsError error = null)
        {
            stopwatch.Stop();
            properties["durationMs"] = stopwatch.ElapsedMilliseconds;
            analyticsService.Track(eventName, "connector", properties, error: error);
        }

        private static Dictionary<string, object> BuildConnectorProperties(string systemKey, string projectId)
        {
            return new Dictionary<string, object>(StringComparer.Ordinal)
            {
                ["systemKey"] = systemKey ?? string.Empty,
                ["projectId"] = projectId ?? string.Empty,
            };
        }

        private static AnalyticsError ToAnalyticsError(Exception ex)
        {
            return new AnalyticsError
            {
                Code = "connector_failed",
                Message = ex.Message,
                ExceptionType = ex.GetType().Name,
            };
        }

        private SheetBinding MergeExistingLayout(SheetBinding bindingSeed)
        {
            if (bindingSeed == null)
            {
                throw new ArgumentNullException(nameof(bindingSeed));
            }

            try
            {
                var existing = metadataStore.LoadBinding(bindingSeed.SheetName);
                return new SheetBinding
                {
                    SheetName = bindingSeed.SheetName,
                    SystemKey = bindingSeed.SystemKey,
                    ProjectId = bindingSeed.ProjectId,
                    ProjectName = bindingSeed.ProjectName,
                    HeaderStartRow = existing.HeaderStartRow > 0 ? existing.HeaderStartRow : bindingSeed.HeaderStartRow,
                    HeaderRowCount = existing.HeaderRowCount > 0 ? existing.HeaderRowCount : bindingSeed.HeaderRowCount,
                    DataStartRow = existing.DataStartRow > 0 ? existing.DataStartRow : bindingSeed.DataStartRow,
                };
            }
            catch (InvalidOperationException)
            {
                return bindingSeed;
            }
        }
    }
}
