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
        public void InitializeSheetSavesBindingAndFieldMappingsFromConnectorSeeds()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var service = CreateService(connector, metadataStore);
            var project = new ProjectOption
            {
                SystemKey = "current-business-system",
                ProjectId = "performance",
                DisplayName = "绩效项目",
            };

            service.InitializeSheet("Sheet1", project);

            Assert.Equal("Sheet1", connector.LastCreateBindingSeedSheetName);
            Assert.Same(project, connector.LastCreateBindingSeedProject);
            Assert.Equal("performance", connector.LastFieldMappingDefinitionProjectId);
            Assert.Equal("Sheet1", connector.LastBuildFieldMappingSeedSheetName);
            Assert.Equal("performance", connector.LastBuildFieldMappingSeedProjectId);
            Assert.Equal("Sheet1", metadataStore.LastSavedBinding.SheetName);
            Assert.Same(connector.FieldMappingDefinition, metadataStore.LastSavedFieldMappingDefinition);
            Assert.Equal(connector.FieldMappingSeedRows.Select(row => row.Values["ApiFieldKey"]), metadataStore.LastSavedFieldMappings.Select(row => row.Values["ApiFieldKey"]));
        }

        [Fact]
        public void InitializeSheetThrowsWhenProjectIsMissing()
        {
            var service = CreateService(new FakeSystemConnector(), new FakeWorksheetMetadataStore());

            Assert.Throws<ArgumentNullException>(() => service.InitializeSheet("Sheet1", null));
        }

        [Fact]
        public void LoadBindingReturnsStoredBinding()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var expected = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
                DataStartRow = 6,
            };
            metadataStore.Bindings["Sheet1"] = expected;
            var service = CreateService(connector, metadataStore);

            var binding = service.LoadBinding("Sheet1");

            Assert.Same(expected, binding);
        }

        [Fact]
        public void LoadFieldMappingsUsesConnectorDefinitionForProject()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            metadataStore.FieldMappings["Sheet1"] = connector.FieldMappingSeedRows.ToArray();
            var service = CreateService(connector, metadataStore);

            var rows = service.LoadFieldMappings("Sheet1", "current-business-system", "performance");

            Assert.Equal("performance", connector.LastFieldMappingDefinitionProjectId);
            Assert.Same(connector.FieldMappingDefinition, metadataStore.LastLoadFieldMappingsDefinition);
            Assert.Equal(connector.FieldMappingSeedRows.Select(row => row.Values["ApiFieldKey"]), rows.Select(row => row.Values["ApiFieldKey"]));
        }

        [Fact]
        public void DownloadDelegatesToConnectorFind()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var service = CreateService(connector, metadataStore);
            var rowIds = new[] { "row-1", "row-2" };
            var fieldKeys = new[] { "owner_name", "start_12345678" };

            var rows = service.Download("current-business-system", "performance", rowIds, fieldKeys);

            Assert.Same(connector.FindResult, rows);
            Assert.Equal("performance", connector.LastFindProjectId);
            Assert.Equal(rowIds, connector.LastFindRowIds);
            Assert.Equal(fieldKeys, connector.LastFindFieldKeys);
        }

        [Fact]
        public void UploadDelegatesToConnectorBatchSave()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var service = CreateService(connector, metadataStore);
            var changes = new[]
            {
                new CellChange
                {
                    SheetName = "Sheet1",
                    RowId = "row-1",
                    ApiFieldKey = "owner_name",
                    NewValue = "李四",
                },
            };

            service.Upload("current-business-system", "performance", changes);

            Assert.Equal("performance", connector.LastBatchSaveProjectId);
            Assert.Same(changes, connector.LastBatchSaveChanges);
        }

        [Fact]
        public void FilterUploadChangesUsesConnectorFilterWhenAvailable()
        {
            var connector = new FakeSystemConnector
            {
                SkippedApiFieldKey = "status",
                SkipReason = "单据已归档，禁止上传",
            };
            var metadataStore = new FakeWorksheetMetadataStore();
            var service = CreateService(connector, metadataStore);
            var changes = new[]
            {
                new CellChange { SheetName = "Sheet1", RowId = "row-1", ApiFieldKey = "owner_name", NewValue = "李四" },
                new CellChange { SheetName = "Sheet1", RowId = "row-1", ApiFieldKey = "status", NewValue = "已完成" },
            };

            var result = service.FilterUploadChanges("current-business-system", "performance", changes);

            Assert.Equal("performance", connector.LastFilterProjectId);
            Assert.Single(result.IncludedChanges);
            Assert.Equal("owner_name", result.IncludedChanges[0].ApiFieldKey);
            var skipped = Assert.Single(result.SkippedChanges);
            Assert.Equal("status", skipped.Change.ApiFieldKey);
            Assert.Equal("单据已归档，禁止上传", skipped.Reason);
        }

        [Fact]
        public void InitializeSheetUsesConnectorMatchingProjectSystemKey()
        {
            var connectorA = new FakeSystemConnector("system-a");
            var connectorB = new FakeSystemConnector("system-b");
            var metadataStore = new FakeWorksheetMetadataStore();
            var service = new WorksheetSyncService(
                new SystemConnectorRegistry(new ISystemConnector[] { connectorA, connectorB }),
                metadataStore,
                new WorksheetChangeTracker(),
                new SyncOperationPreviewFactory());

            service.InitializeSheet(
                "Sheet1",
                new ProjectOption
                {
                    SystemKey = "system-b",
                    ProjectId = "shared-project",
                    DisplayName = "项目B",
                });

            Assert.Null(connectorA.LastCreateBindingSeedProject);
            Assert.NotNull(connectorB.LastCreateBindingSeedProject);
            Assert.Equal("system-b", metadataStore.LastSavedBinding.SystemKey);
        }

        private static WorksheetSyncService CreateService(
            FakeSystemConnector connector,
            FakeWorksheetMetadataStore metadataStore)
        {
            return new WorksheetSyncService(
                new SystemConnectorRegistry(new[] { connector }),
                metadataStore,
                new WorksheetChangeTracker(),
                new SyncOperationPreviewFactory());
        }

        private sealed class FakeSystemConnector : ISystemConnector, IUploadChangeFilter
        {
            public FakeSystemConnector(string systemKey = "current-business-system")
            {
                SystemKey = systemKey;
                FieldMappingDefinition = new FieldMappingTableDefinition
                {
                    SystemKey = systemKey,
                    Columns = new[]
                    {
                        new FieldMappingColumnDefinition { ColumnName = "ApiFieldKey", Role = FieldMappingSemanticRole.ApiFieldKey },
                        new FieldMappingColumnDefinition { ColumnName = "HeaderType", Role = FieldMappingSemanticRole.HeaderType },
                        new FieldMappingColumnDefinition { ColumnName = "IsIdColumn", Role = FieldMappingSemanticRole.IsIdColumn },
                        new FieldMappingColumnDefinition { ColumnName = "CurrentSingleDisplayName", Role = FieldMappingSemanticRole.CurrentSingleHeaderText },
                        new FieldMappingColumnDefinition { ColumnName = "CurrentParentDisplayName", Role = FieldMappingSemanticRole.CurrentParentHeaderText },
                        new FieldMappingColumnDefinition { ColumnName = "CurrentChildDisplayName", Role = FieldMappingSemanticRole.CurrentChildHeaderText },
                    },
                };

                FieldMappingSeedRows = new[]
                {
                    CreateMappingRow("Sheet1", "row_id", "single", true, currentSingle: "ID"),
                    CreateMappingRow("Sheet1", "owner_name", "single", false, currentSingle: "项目负责人"),
                    CreateMappingRow("Sheet1", "start_12345678", "activityProperty", false, currentParent: "测试活动111", currentChild: "开始时间"),
                };
            }

            public string SystemKey { get; }

            public SheetBinding BindingSeed { get; set; } = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 1,
                HeaderRowCount = 2,
                DataStartRow = 3,
            };

            public FieldMappingTableDefinition FieldMappingDefinition { get; }

            public IReadOnlyList<SheetFieldMappingRow> FieldMappingSeedRows { get; }

            public IReadOnlyList<IDictionary<string, object>> FindResult { get; } = new[]
            {
                (IDictionary<string, object>)new Dictionary<string, object>(StringComparer.Ordinal)
                {
                    ["row_id"] = "row-1",
                    ["owner_name"] = "张三",
                },
            };

            public string LastCreateBindingSeedSheetName { get; private set; }

            public ProjectOption LastCreateBindingSeedProject { get; private set; }

            public string LastFieldMappingDefinitionProjectId { get; private set; }

            public string LastBuildFieldMappingSeedSheetName { get; private set; }

            public string LastBuildFieldMappingSeedProjectId { get; private set; }

            public string LastFindProjectId { get; private set; }

            public IReadOnlyList<string> LastFindRowIds { get; private set; } = Array.Empty<string>();

            public IReadOnlyList<string> LastFindFieldKeys { get; private set; } = Array.Empty<string>();

            public string LastBatchSaveProjectId { get; private set; }

            public IReadOnlyList<CellChange> LastBatchSaveChanges { get; private set; } = Array.Empty<CellChange>();

            public string SkippedApiFieldKey { get; set; } = string.Empty;

            public string SkipReason { get; set; } = string.Empty;

            public string LastFilterProjectId { get; private set; }

            public IReadOnlyList<ProjectOption> GetProjects()
            {
                return Array.Empty<ProjectOption>();
            }

            public SheetBinding CreateBindingSeed(string sheetName, ProjectOption project)
            {
                LastCreateBindingSeedSheetName = sheetName;
                LastCreateBindingSeedProject = project;
                return new SheetBinding
                {
                    SheetName = sheetName,
                    SystemKey = project?.SystemKey ?? SystemKey,
                    ProjectId = project?.ProjectId ?? BindingSeed.ProjectId,
                    ProjectName = project?.DisplayName ?? BindingSeed.ProjectName,
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
                LastBuildFieldMappingSeedSheetName = sheetName;
                LastBuildFieldMappingSeedProjectId = projectId;
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
                LastFindRowIds = rowIds;
                LastFindFieldKeys = fieldKeys;
                return FindResult;
            }

            public void BatchSave(string projectId, IReadOnlyList<CellChange> changes)
            {
                LastBatchSaveProjectId = projectId;
                LastBatchSaveChanges = changes;
            }

            public UploadChangeFilterResult FilterUploadChanges(string projectId, IReadOnlyList<CellChange> changes)
            {
                LastFilterProjectId = projectId;

                return new UploadChangeFilterResult
                {
                    IncludedChanges = (changes ?? Array.Empty<CellChange>())
                        .Where(change => !string.Equals(change.ApiFieldKey, SkippedApiFieldKey, StringComparison.Ordinal))
                        .ToArray(),
                    SkippedChanges = (changes ?? Array.Empty<CellChange>())
                        .Where(change => string.Equals(change.ApiFieldKey, SkippedApiFieldKey, StringComparison.Ordinal))
                        .Select(change => new SkippedCellChange
                        {
                            Change = change,
                            Reason = SkipReason,
                        })
                        .ToArray(),
                };
            }

            private static SheetFieldMappingRow CreateMappingRow(
                string sheetName,
                string apiFieldKey,
                string headerType,
                bool isIdColumn,
                string currentSingle = "",
                string currentParent = "",
                string currentChild = "")
            {
                return new SheetFieldMappingRow
                {
                    SheetName = sheetName,
                    Values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["ApiFieldKey"] = apiFieldKey,
                        ["HeaderType"] = headerType,
                        ["IsIdColumn"] = isIdColumn ? "true" : "false",
                        ["CurrentSingleDisplayName"] = currentSingle,
                        ["CurrentParentDisplayName"] = currentParent,
                        ["CurrentChildDisplayName"] = currentChild,
                    },
                };
            }
        }

        private sealed class FakeWorksheetMetadataStore : IWorksheetMetadataStore
        {
            public Dictionary<string, SheetBinding> Bindings { get; } = new Dictionary<string, SheetBinding>(StringComparer.OrdinalIgnoreCase);

            public Dictionary<string, SheetFieldMappingRow[]> FieldMappings { get; } = new Dictionary<string, SheetFieldMappingRow[]>(StringComparer.OrdinalIgnoreCase);

            public SheetBinding LastSavedBinding { get; private set; }

            public FieldMappingTableDefinition LastSavedFieldMappingDefinition { get; private set; }

            public SheetFieldMappingRow[] LastSavedFieldMappings { get; private set; } = Array.Empty<SheetFieldMappingRow>();

            public FieldMappingTableDefinition LastLoadFieldMappingsDefinition { get; private set; }

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
                LastLoadFieldMappingsDefinition = definition;
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
    }
}
