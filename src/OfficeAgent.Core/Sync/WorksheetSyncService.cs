using System;
using System.Collections.Generic;
using System.Linq;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;

namespace OfficeAgent.Core.Sync
{
    public sealed class WorksheetSyncService
    {
        private readonly ISystemConnectorRegistry connectorRegistry;
        private readonly IWorksheetMetadataStore metadataStore;

        public WorksheetSyncService(
            ISystemConnectorRegistry connectorRegistry,
            IWorksheetMetadataStore metadataStore)
        {
            this.connectorRegistry = connectorRegistry ?? throw new ArgumentNullException(nameof(connectorRegistry));
            this.metadataStore = metadataStore ?? throw new ArgumentNullException(nameof(metadataStore));
        }

        public WorksheetSyncService(
            ISystemConnectorRegistry connectorRegistry,
            IWorksheetMetadataStore metadataStore,
            WorksheetChangeTracker changeTracker,
            SyncOperationPreviewFactory previewFactory)
            : this(connectorRegistry, metadataStore)
        {
        }

        public IReadOnlyList<ProjectOption> GetProjects()
        {
            return connectorRegistry.GetProjects() ?? Array.Empty<ProjectOption>();
        }

        public void InitializeSheet(string sheetName, ProjectOption project)
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
        }

        public SheetBinding CreateBindingSeed(string sheetName, ProjectOption project)
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
            return MergeExistingLayout(connector.CreateBindingSeed(sheetName, project));
        }

        public SheetBinding LoadBinding(string sheetName)
        {
            return metadataStore.LoadBinding(sheetName);
        }

        public FieldMappingTableDefinition LoadFieldMappingDefinition(string systemKey, string projectId)
        {
            return GetRequiredConnector(systemKey).GetFieldMappingDefinition(projectId);
        }

        public SheetFieldMappingRow[] LoadFieldMappings(string sheetName, string systemKey, string projectId)
        {
            var definition = LoadFieldMappingDefinition(systemKey, projectId);
            return metadataStore.LoadFieldMappings(sheetName, definition);
        }

        public IReadOnlyList<IDictionary<string, object>> Download(
            string systemKey,
            string projectId,
            IReadOnlyList<string> rowIds,
            IReadOnlyList<string> fieldKeys)
        {
            return GetRequiredConnector(systemKey).Find(projectId, rowIds, fieldKeys);
        }

        public void Upload(string systemKey, string projectId, IReadOnlyList<CellChange> changes)
        {
            GetRequiredConnector(systemKey).BatchSave(projectId, changes);
        }

        public UploadChangeFilterResult FilterUploadChanges(
            string systemKey,
            string projectId,
            IReadOnlyList<CellChange> changes)
        {
            var changeList = changes ?? Array.Empty<CellChange>();
            var connector = GetRequiredConnector(systemKey);
            if (connector is IUploadChangeFilter filter)
            {
                var result = filter.FilterUploadChanges(projectId, changeList);
                if (result != null)
                {
                    return new UploadChangeFilterResult
                    {
                        IncludedChanges = result.IncludedChanges ?? Array.Empty<CellChange>(),
                        SkippedChanges = result.SkippedChanges ?? Array.Empty<SkippedCellChange>(),
                    };
                }
            }

            return new UploadChangeFilterResult
            {
                IncludedChanges = changeList.ToArray(),
                SkippedChanges = Array.Empty<SkippedCellChange>(),
            };
        }

        private ISystemConnector GetRequiredConnector(string systemKey)
        {
            return connectorRegistry.GetRequiredConnector(systemKey);
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
