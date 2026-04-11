using System;
using System.Collections.Generic;
using System.Linq;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
using OfficeAgent.Core.Sync;
using Xunit;

namespace OfficeAgent.Core.Tests
{
    public sealed class WorksheetSyncServiceTests
    {
        [Fact]
        public void PrepareIncrementalUploadBuildsPreviewFromDirtySnapshotBackedCells()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            metadataStore.SnapshotsToReturn.Enqueue(new[]
            {
                new WorksheetSnapshotCell { SheetName = "Sync-performance", RowId = "row-1", ApiFieldKey = "name", Value = "old name" },
                new WorksheetSnapshotCell { SheetName = "Sync-performance", RowId = "row-2", ApiFieldKey = "status", Value = "inactive" },
            });

            var service = new WorksheetSyncService(
                connector,
                metadataStore,
                new WorksheetChangeTracker(),
                new SyncOperationPreviewFactory());

            var preview = service.PrepareIncrementalUpload(
                "Sync-performance",
                new[]
                {
                    new CellChange { SheetName = "Sync-performance", RowId = "row-1", ApiFieldKey = "name", OldValue = "old name", NewValue = "new name" },
                    new CellChange { SheetName = "Sync-performance", RowId = "row-2", ApiFieldKey = "status", OldValue = "inactive", NewValue = "inactive" },
                    new CellChange { SheetName = "Sync-performance", RowId = string.Empty, ApiFieldKey = "name", OldValue = string.Empty, NewValue = "ignored" },
                });

            Assert.Equal("增量上传", preview.OperationName);
            var changed = Assert.Single(preview.Changes);
            Assert.Equal("row-1", changed.RowId);
            Assert.Equal("name", changed.ApiFieldKey);
            Assert.Equal("new name", changed.NewValue);
        }

        [Fact]
        public void PrepareIncrementalUploadTreatsMissingSnapshotAsEmpty()
        {
            var service = CreateService(out _, out var metadataStore);
            metadataStore.SnapshotsToReturn.Enqueue(null);

            var preview = service.PrepareIncrementalUpload(
                "Sync-performance",
                new[]
                {
                    new CellChange { SheetName = "Sync-performance", RowId = "row-1", ApiFieldKey = "name", OldValue = "old", NewValue = "new" },
                });

            Assert.Empty(preview.Changes);
        }

        [Fact]
        public void PrepareIncrementalUploadReloadsLatestSnapshotOnEachCall()
        {
            var service = CreateService(out _, out var metadataStore);
            metadataStore.SnapshotsToReturn.Enqueue(new[]
            {
                new WorksheetSnapshotCell { SheetName = "Sync-performance", RowId = "row-1", ApiFieldKey = "name", Value = "v1" },
            });
            metadataStore.SnapshotsToReturn.Enqueue(new[]
            {
                new WorksheetSnapshotCell { SheetName = "Sync-performance", RowId = "row-1", ApiFieldKey = "name", Value = "v2" },
            });

            var first = service.PrepareIncrementalUpload(
                "Sync-performance",
                new[]
                {
                    new CellChange { SheetName = "Sync-performance", RowId = "row-1", ApiFieldKey = "name", OldValue = "v1", NewValue = "v2" },
                });

            var second = service.PrepareIncrementalUpload(
                "Sync-performance",
                new[]
                {
                    new CellChange { SheetName = "Sync-performance", RowId = "row-1", ApiFieldKey = "name", OldValue = "v1", NewValue = "v2" },
                });

            Assert.Single(first.Changes);
            Assert.Empty(second.Changes);
            Assert.Equal(2, metadataStore.LoadSnapshotCalls);
        }

        [Fact]
        public void LoadSchemaForSheetUsesLatestBindingOnEachCall()
        {
            var service = CreateService(out var connector, out var metadataStore);
            var firstSchema = new WorksheetSchema { ProjectId = "project-1" };
            var secondSchema = new WorksheetSchema { ProjectId = "project-2" };
            metadataStore.BindingsToReturn.Enqueue(new SheetBinding { SheetName = "Sync", ProjectId = "project-1" });
            metadataStore.BindingsToReturn.Enqueue(new SheetBinding { SheetName = "Sync", ProjectId = "project-2" });
            connector.SchemaByProjectId["project-1"] = firstSchema;
            connector.SchemaByProjectId["project-2"] = secondSchema;

            var schema1 = service.LoadSchemaForSheet("Sync");
            var schema2 = service.LoadSchemaForSheet("Sync");

            Assert.Same(firstSchema, schema1);
            Assert.Same(secondSchema, schema2);
            Assert.Equal(new[] { "project-1", "project-2" }, connector.GetSchemaProjectIds);
            Assert.Equal(2, metadataStore.LoadBindingCalls);
        }

        [Fact]
        public void LoadSchemaForSheetLetsMissingBindingErrorSurface()
        {
            var service = CreateService(out var connector, out var metadataStore);
            metadataStore.BindingLoadError = new InvalidOperationException("Missing binding");

            var error = Assert.Throws<InvalidOperationException>(() => service.LoadSchemaForSheet("Sync"));

            Assert.Equal("Missing binding", error.Message);
            Assert.Empty(connector.GetSchemaProjectIds);
        }

        [Fact]
        public void ExecutePartialDownloadLoadsBindingAndCallsFind()
        {
            var service = CreateService(out var connector, out var metadataStore);
            metadataStore.BindingsToReturn.Enqueue(new SheetBinding { SheetName = "Sync", ProjectId = "project-9" });
            var selection = new ResolvedSelection
            {
                RowIds = new[] { "row-1", "row-2" },
                ApiFieldKeys = new[] { "name", "status" },
            };
            var findResult = new[]
            {
                new Dictionary<string, object> { ["rowId"] = "row-1", ["name"] = "Alpha" },
            };
            connector.FindResult = findResult;

            var result = service.ExecutePartialDownload("Sync", selection);

            Assert.Same(findResult, result);
            Assert.Equal("project-9", connector.LastFindProjectId);
            Assert.Equal(selection.RowIds, connector.LastFindRowIds);
            Assert.Equal(selection.ApiFieldKeys, connector.LastFindFieldKeys);
        }

        [Fact]
        public void ExecutePartialUploadLoadsBindingAndCallsBatchSave()
        {
            var service = CreateService(out var connector, out var metadataStore);
            metadataStore.BindingsToReturn.Enqueue(new SheetBinding { SheetName = "Sync", ProjectId = "project-77" });
            var changes = new[]
            {
                new CellChange { SheetName = "Sync", RowId = "row-1", ApiFieldKey = "name", NewValue = "Next" },
            };

            service.ExecutePartialUpload("Sync", changes);

            Assert.Equal("project-77", connector.LastBatchSaveProjectId);
            Assert.Same(changes, connector.LastBatchSaveChanges);
        }

        [Fact]
        public void ExecutePartialUploadReloadsLatestBindingOnEachCall()
        {
            var service = CreateService(out var connector, out var metadataStore);
            metadataStore.BindingsToReturn.Enqueue(new SheetBinding { SheetName = "Sync", ProjectId = "project-1" });
            metadataStore.BindingsToReturn.Enqueue(new SheetBinding { SheetName = "Sync", ProjectId = "project-2" });
            var changes = new[] { new CellChange { SheetName = "Sync", RowId = "row-1", ApiFieldKey = "name", NewValue = "A" } };

            service.ExecutePartialUpload("Sync", changes);
            service.ExecutePartialUpload("Sync", changes);

            Assert.Equal(new[] { "project-1", "project-2" }, connector.BatchSaveProjectIds);
            Assert.Equal(2, metadataStore.LoadBindingCalls);
        }

        private static WorksheetSyncService CreateService(out FakeSystemConnector connector, out FakeWorksheetMetadataStore metadataStore)
        {
            connector = new FakeSystemConnector();
            metadataStore = new FakeWorksheetMetadataStore();

            return new WorksheetSyncService(
                connector,
                metadataStore,
                new WorksheetChangeTracker(),
                new SyncOperationPreviewFactory());
        }

        private sealed class FakeSystemConnector : ISystemConnector
        {
            public Dictionary<string, WorksheetSchema> SchemaByProjectId { get; } = new Dictionary<string, WorksheetSchema>(StringComparer.Ordinal);

            public List<string> GetSchemaProjectIds { get; } = new List<string>();

            public IReadOnlyList<IDictionary<string, object>> FindResult { get; set; } = Array.Empty<IDictionary<string, object>>();

            public string LastFindProjectId { get; private set; }

            public IReadOnlyList<string> LastFindRowIds { get; private set; }

            public IReadOnlyList<string> LastFindFieldKeys { get; private set; }

            public string LastBatchSaveProjectId { get; private set; }

            public IReadOnlyList<CellChange> LastBatchSaveChanges { get; private set; }

            public List<string> BatchSaveProjectIds { get; } = new List<string>();

            public IReadOnlyList<ProjectOption> GetProjects()
            {
                return Array.Empty<ProjectOption>();
            }

            public WorksheetSchema GetSchema(string projectId)
            {
                GetSchemaProjectIds.Add(projectId);
                return SchemaByProjectId[projectId];
            }

            public IReadOnlyList<IDictionary<string, object>> Find(
                string projectId,
                IReadOnlyList<string> rowIds,
                IReadOnlyList<string> fieldKeys)
            {
                LastFindProjectId = projectId;
                LastFindRowIds = rowIds;
                LastFindFieldKeys = fieldKeys;
                return FindResult;
            }

            public void BatchSave(string projectId, IReadOnlyList<CellChange> changes)
            {
                LastBatchSaveProjectId = projectId;
                LastBatchSaveChanges = changes;
                BatchSaveProjectIds.Add(projectId);
            }
        }

        private sealed class FakeWorksheetMetadataStore : IWorksheetMetadataStore
        {
            public Queue<SheetBinding> BindingsToReturn { get; } = new Queue<SheetBinding>();

            public Queue<WorksheetSnapshotCell[]> SnapshotsToReturn { get; } = new Queue<WorksheetSnapshotCell[]>();

            public Exception BindingLoadError { get; set; }

            public int LoadBindingCalls { get; private set; }

            public int LoadSnapshotCalls { get; private set; }

            public void SaveBinding(SheetBinding binding)
            {
                BindingsToReturn.Enqueue(binding);
            }

            public SheetBinding LoadBinding(string sheetName)
            {
                LoadBindingCalls++;
                if (BindingLoadError != null)
                {
                    throw BindingLoadError;
                }

                if (BindingsToReturn.Count == 0)
                {
                    throw new InvalidOperationException("No binding queued.");
                }

                return BindingsToReturn.Dequeue();
            }

            public WorksheetSnapshotCell[] LoadSnapshot(string sheetName)
            {
                LoadSnapshotCalls++;
                if (SnapshotsToReturn.Count == 0)
                {
                    return Array.Empty<WorksheetSnapshotCell>();
                }

                return SnapshotsToReturn.Dequeue();
            }

            public void SaveSnapshot(string sheetName, WorksheetSnapshotCell[] cells)
            {
                SnapshotsToReturn.Enqueue(cells);
            }
        }
    }
}
