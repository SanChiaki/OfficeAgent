using System;
using System.Collections.Generic;
using System.Linq;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;

namespace OfficeAgent.Core.Sync
{
    public sealed class WorksheetSyncService
    {
        private readonly ISystemConnector connector;
        private readonly IWorksheetMetadataStore metadataStore;
        private readonly WorksheetChangeTracker changeTracker;
        private readonly SyncOperationPreviewFactory previewFactory;

        public WorksheetSyncService(
            ISystemConnector connector,
            IWorksheetMetadataStore metadataStore,
            WorksheetChangeTracker changeTracker,
            SyncOperationPreviewFactory previewFactory)
        {
            this.connector = connector ?? throw new ArgumentNullException(nameof(connector));
            this.metadataStore = metadataStore ?? throw new ArgumentNullException(nameof(metadataStore));
            this.changeTracker = changeTracker ?? throw new ArgumentNullException(nameof(changeTracker));
            this.previewFactory = previewFactory ?? throw new ArgumentNullException(nameof(previewFactory));
        }

        public SyncOperationPreview PrepareIncrementalUpload(string sheetName, IReadOnlyList<CellChange> currentCells)
        {
            var snapshot = metadataStore.LoadSnapshot(sheetName) ?? Array.Empty<WorksheetSnapshotCell>();
            var dirtyCells = changeTracker
                .GetDirtyCells(sheetName, snapshot, currentCells ?? Array.Empty<CellChange>())
                .Where(item => !string.IsNullOrWhiteSpace(item.RowId))
                .ToArray();

            return previewFactory.CreateUploadPreview("增量上传", dirtyCells);
        }

        public WorksheetSchema LoadSchemaForSheet(string sheetName)
        {
            var binding = metadataStore.LoadBinding(sheetName);
            return connector.GetSchema(binding.ProjectId);
        }

        public IReadOnlyList<IDictionary<string, object>> ExecutePartialDownload(string sheetName, ResolvedSelection selection)
        {
            var binding = metadataStore.LoadBinding(sheetName);
            return connector.Find(binding.ProjectId, selection.RowIds, selection.ApiFieldKeys);
        }

        public void ExecutePartialUpload(string sheetName, IReadOnlyList<CellChange> changes)
        {
            var binding = metadataStore.LoadBinding(sheetName);
            connector.BatchSave(binding.ProjectId, changes);
        }
    }
}
