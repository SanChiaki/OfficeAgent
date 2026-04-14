using System;
using System.Collections.Generic;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;

namespace OfficeAgent.Core.Sync
{
    public sealed class WorksheetSyncService
    {
        private readonly ISystemConnector connector;
        private readonly IWorksheetMetadataStore metadataStore;

        public WorksheetSyncService(
            ISystemConnector connector,
            IWorksheetMetadataStore metadataStore)
        {
            this.connector = connector ?? throw new ArgumentNullException(nameof(connector));
            this.metadataStore = metadataStore ?? throw new ArgumentNullException(nameof(metadataStore));
        }

        public WorksheetSyncService(
            ISystemConnector connector,
            IWorksheetMetadataStore metadataStore,
            WorksheetChangeTracker changeTracker,
            SyncOperationPreviewFactory previewFactory)
            : this(connector, metadataStore)
        {
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

            var bindingSeed = connector.CreateBindingSeed(sheetName, project);
            var binding = MergeExistingLayout(bindingSeed);
            var definition = connector.GetFieldMappingDefinition(project.ProjectId);
            var seedRows = connector.BuildFieldMappingSeed(sheetName, project.ProjectId);

            metadataStore.SaveBinding(binding);
            metadataStore.SaveFieldMappings(sheetName, definition, seedRows);
        }

        public SheetBinding LoadBinding(string sheetName)
        {
            return metadataStore.LoadBinding(sheetName);
        }

        public FieldMappingTableDefinition LoadFieldMappingDefinition(string projectId)
        {
            return connector.GetFieldMappingDefinition(projectId);
        }

        public SheetFieldMappingRow[] LoadFieldMappings(string sheetName, string projectId)
        {
            var definition = connector.GetFieldMappingDefinition(projectId);
            return metadataStore.LoadFieldMappings(sheetName, definition);
        }

        public IReadOnlyList<IDictionary<string, object>> Download(
            string projectId,
            IReadOnlyList<string> rowIds,
            IReadOnlyList<string> fieldKeys)
        {
            return connector.Find(projectId, rowIds, fieldKeys);
        }

        public void Upload(string projectId, IReadOnlyList<CellChange> changes)
        {
            connector.BatchSave(projectId, changes);
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
