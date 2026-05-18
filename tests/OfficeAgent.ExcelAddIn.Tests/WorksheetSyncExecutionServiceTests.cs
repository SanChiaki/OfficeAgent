using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using System.Threading;
using System.Threading.Tasks;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
using OfficeAgent.Core.Sync;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class WorksheetSyncExecutionServiceTests
    {
        [Fact]
        public void InitializeCurrentSheetWritesBindingAndFieldMappingsWithoutTouchingBusinessCells()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var selectionReader = new FakeWorksheetSelectionReader();
            var (service, grid) = CreateService(connector, metadataStore, selectionReader);

            grid.SetCell("Sheet1", 1, 1, "现有说明");

            InvokeInitialize(service, "Sheet1", new ProjectOption
            {
                SystemKey = "current-business-system",
                ProjectId = "performance",
                DisplayName = "绩效项目",
            });

            Assert.Equal("现有说明", grid.GetCell("Sheet1", 1, 1));
            Assert.Equal(1, metadataStore.LastSavedBinding.HeaderStartRow);
            Assert.Equal("performance", connector.LastFieldMappingDefinitionProjectId);
            Assert.NotEmpty(metadataStore.LastSavedFieldMappings);
        }

        [Fact]
        public void TryAutoInitializeCurrentSheetReinitializesWhenSystemKeyChangesButProjectIdMatches()
        {
            var connectorA = new FakeSystemConnector("system-a");
            var connectorB = new FakeSystemConnector("system-b");
            var metadataStore = new FakeWorksheetMetadataStore();
            metadataStore.Bindings["Sheet1"] = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "system-a",
                ProjectId = "shared-project",
                ProjectName = "旧项目",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
                DataStartRow = 6,
            };
            metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");

            var (service, _) = CreateService(
                new[] { connectorA, connectorB },
                metadataStore,
                new FakeWorksheetSelectionReader());

            InvokeTryAutoInitialize(service, "Sheet1", new ProjectOption
            {
                SystemKey = "system-b",
                ProjectId = "shared-project",
                DisplayName = "新项目",
            });

            Assert.Equal("system-b", metadataStore.LastSavedBinding.SystemKey);
            Assert.Equal("shared-project", metadataStore.LastSavedBinding.ProjectId);
            Assert.Equal("新项目", metadataStore.LastSavedBinding.ProjectName);
            Assert.Null(connectorA.LastCreateBindingSeedProject);
            Assert.NotNull(connectorB.LastCreateBindingSeedProject);
        }

        [Fact]
        public void ExecuteFullDownloadHonorsConfiguredHeaderAndDataRowsWhenSheetHeadersAreEmpty()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            metadataStore.Bindings["Sheet1"] = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
                DataStartRow = 6,
            };
            metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");
            connector.FindResult = new[] { CreateRow("row-1", "张三", "2026-01-02", "2026-01-05") };

            var (service, grid) = CreateService(connector, metadataStore, new FakeWorksheetSelectionReader());
            grid.SetCell("Sheet1", 1, 1, "统计说明");
            grid.SetCell("Sheet1", 5, 1, "统计行");

            var plan = InvokePrepare(service, "PrepareFullDownload", "Sheet1");
            InvokeExecute(service, "ExecuteDownload", plan);

            Assert.Equal("统计说明", grid.GetCell("Sheet1", 1, 1));
            Assert.Equal("统计行", grid.GetCell("Sheet1", 5, 1));
            Assert.Equal("ID", grid.GetCell("Sheet1", 3, 1));
            Assert.Equal("项目负责人", grid.GetCell("Sheet1", 3, 2));
            Assert.Equal("测试活动111", grid.GetCell("Sheet1", 3, 3));
            Assert.Equal("开始时间", grid.GetCell("Sheet1", 4, 3));
            Assert.Equal("结束时间", grid.GetCell("Sheet1", 4, 4));
            Assert.Equal("row-1", grid.GetCell("Sheet1", 6, 1));
            Assert.Equal("张三", grid.GetCell("Sheet1", 6, 2));
            Assert.Equal("2026-01-02", grid.GetCell("Sheet1", 6, 3));
            Assert.Equal("2026-01-05", grid.GetCell("Sheet1", 6, 4));

            Assert.Contains(grid.Merges, merge => merge.SheetName == "Sheet1" && merge.Row == 3 && merge.Column == 1 && merge.RowSpan == 2 && merge.ColumnSpan == 1);
            Assert.Contains(grid.Merges, merge => merge.SheetName == "Sheet1" && merge.Row == 3 && merge.Column == 3 && merge.RowSpan == 1 && merge.ColumnSpan == 2);
        }

        [Fact]
        public void ExecuteFullDownloadUsesActivityExcelL1WhenSingleRowSheetHeadersAreEmpty()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            metadataStore.Bindings["Sheet1"] = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 1,
                HeaderRowCount = 1,
                DataStartRow = 2,
            };
            metadataStore.FieldMappings["Sheet1"] = BuildSingleRowActivityL1Mappings("Sheet1");
            connector.FindResult = new[] { CreateRow("row-1", "张三", "2026-01-02", "2026-01-05") };

            var (service, grid) = CreateService(connector, metadataStore, new FakeWorksheetSelectionReader());

            var plan = InvokePrepare(service, "PrepareFullDownload", "Sheet1");
            InvokeExecute(service, "ExecuteDownload", plan);

            Assert.Equal("ID", grid.GetCell("Sheet1", 1, 1));
            Assert.Equal("项目负责人", grid.GetCell("Sheet1", 1, 2));
            Assert.Equal("计划开始", grid.GetCell("Sheet1", 1, 3));
            Assert.Equal("计划结束", grid.GetCell("Sheet1", 1, 4));
            Assert.Equal("2026-01-02", grid.GetCell("Sheet1", 2, 3));
            Assert.Equal("2026-01-05", grid.GetCell("Sheet1", 2, 4));
            Assert.DoesNotContain(grid.Merges, merge => merge.RowSpan > 1 || merge.ColumnSpan > 1);
        }

        [Fact]
        public void ExecutePartialDownloadMatchesActivityExcelL1ForSingleRowHeaders()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 1,
                HeaderRowCount = 1,
                DataStartRow = 2,
            };
            metadataStore.Bindings["Sheet1"] = binding;
            metadataStore.FieldMappings["Sheet1"] = BuildSingleRowActivityL1Mappings("Sheet1");
            connector.FindResult = new[] { CreateRow("row-1", "张三", "2026-02-01", "2026-02-09") };

            var selectionReader = new FakeWorksheetSelectionReader
            {
                VisibleCells = new[]
                {
                    new SelectedVisibleCell { Row = 2, Column = 3, Value = "旧开始时间" },
                },
            };
            var (service, grid) = CreateService(connector, metadataStore, selectionReader);
            grid.SetCell("Sheet1", 1, 1, "ID");
            grid.SetCell("Sheet1", 1, 2, "项目负责人");
            grid.SetCell("Sheet1", 1, 3, "计划开始");
            grid.SetCell("Sheet1", 1, 4, "计划结束");
            grid.SetCell("Sheet1", 2, 1, "row-1");
            grid.SetCell("Sheet1", 2, 3, "旧开始时间");

            var plan = InvokePrepare(service, "PreparePartialDownload", "Sheet1");
            InvokeExecute(service, "ExecuteDownload", plan);

            Assert.Equal(new[] { "row-1" }, connector.LastFindRowIds);
            Assert.Equal(new[] { "start_12345678" }, connector.LastFindFieldKeys);
            Assert.Equal("2026-02-01", grid.GetCell("Sheet1", 2, 3));
        }

        [Fact]
        public void ExecutePartialDownloadUsesRecognizedHeadersAndIdLookupOutsideSelection()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
                DataStartRow = 6,
            };
            metadataStore.Bindings["Sheet1"] = binding;
            metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");
            connector.FindResult = new[] { CreateRow("row-1", "张三", "2026-02-01", "2026-02-09") };

            var selectionReader = new FakeWorksheetSelectionReader
            {
                VisibleCells = new[]
                {
                    new SelectedVisibleCell { Row = 6, Column = 3, Value = "旧开始时间" },
                },
            };
            var (service, grid) = CreateService(connector, metadataStore, selectionReader);
            SeedRecognizedHeaders(grid, "Sheet1", binding);
            grid.SetCell("Sheet1", 6, 1, "row-1");
            grid.SetCell("Sheet1", 6, 2, "张三");
            grid.SetCell("Sheet1", 6, 3, "旧开始时间");
            grid.SetCell("Sheet1", 6, 4, "旧结束时间");

            var plan = InvokePrepare(service, "PreparePartialDownload", "Sheet1");
            InvokeExecute(service, "ExecuteDownload", plan);

            Assert.Equal("performance", connector.LastFindProjectId);
            Assert.Equal(new[] { "row-1" }, connector.LastFindRowIds);
            Assert.Equal(new[] { "start_12345678" }, connector.LastFindFieldKeys);
            Assert.Equal("2026-02-01", grid.GetCell("Sheet1", 6, 3));
            Assert.Equal("旧结束时间", grid.GetCell("Sheet1", 6, 4));
        }

        [Fact]
        public void ExecutePartialDownloadAppendsWorkbookLogForChangedCells()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
                DataStartRow = 6,
            };
            metadataStore.Bindings["Sheet1"] = binding;
            metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");
            connector.FindResult = new[] { CreateRow("row-1", "张三", "2026-02-01", "2026-02-09") };

            var selectionReader = new FakeWorksheetSelectionReader
            {
                VisibleCells = new[]
                {
                    new SelectedVisibleCell { Row = 6, Column = 3, Value = "旧开始时间" },
                },
            };
            var (service, grid, logStore, _) = CreateServiceWithChangeLog(connector, metadataStore, selectionReader);
            SeedRecognizedHeaders(grid, "Sheet1", binding);
            grid.SetCell("Sheet1", 6, 1, "row-1");
            grid.SetCell("Sheet1", 6, 3, "旧开始时间");

            var plan = InvokePrepare(service, "PreparePartialDownload", "Sheet1");
            InvokeExecute(service, "ExecuteDownload", plan);

            var entry = Assert.Single(logStore.Entries);
            Assert.Equal("row-1", entry.Key);
            Assert.Equal("测试活动111/开始时间", entry.HeaderText);
            Assert.Equal("Download", entry.ChangeMode);
            Assert.Equal("2026-02-01", entry.NewValue);
            Assert.Equal("旧开始时间", entry.OldValue);
        }

        [Fact]
        public void PreparePartialDownloadResolvesGroupedSingleOwnerNameFromTwoRowHeader()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
                DataStartRow = 6,
            };
            metadataStore.Bindings["Sheet1"] = binding;
            metadataStore.FieldMappings["Sheet1"] = BuildGroupedSingleOwnerMappings("Sheet1");
            connector.FindResult = new[] { CreateRow("row-1", "张三", "2026-02-01", "2026-02-09") };

            var selectionReader = new FakeWorksheetSelectionReader
            {
                VisibleCells = new[]
                {
                    new SelectedVisibleCell { Row = 6, Column = 2, Value = "旧负责人" },
                },
            };
            var (service, grid) = CreateService(connector, metadataStore, selectionReader);
            SeedGroupedSingleRecognizedHeaders(grid, "Sheet1", binding);
            grid.SetCell("Sheet1", 6, 1, "row-1");
            grid.SetCell("Sheet1", 6, 2, "旧负责人");

            _ = InvokePrepare(service, "PreparePartialDownload", "Sheet1");

            Assert.Equal(new[] { "row-1" }, connector.LastFindRowIds);
            Assert.Equal(new[] { "owner_name" }, connector.LastFindFieldKeys);
        }

        [Fact]
        public void ExecutePartialDownloadBatchesRectangularSelectionWrites()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
                DataStartRow = 6,
            };
            metadataStore.Bindings["Sheet1"] = binding;
            metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");
            connector.FindResult = new[]
            {
                CreateRow("row-1", "张三", "2026-02-01", "2026-02-09"),
                CreateRow("row-2", "李四", "2026-03-01", "2026-03-09"),
            };

            var selectionReader = new FakeWorksheetSelectionReader
            {
                VisibleCells = new[]
                {
                    new SelectedVisibleCell { Row = 6, Column = 3, Value = "旧开始时间1" },
                    new SelectedVisibleCell { Row = 6, Column = 4, Value = "旧结束时间1" },
                    new SelectedVisibleCell { Row = 7, Column = 3, Value = "旧开始时间2" },
                    new SelectedVisibleCell { Row = 7, Column = 4, Value = "旧结束时间2" },
                },
            };
            var (service, grid) = CreateService(connector, metadataStore, selectionReader);
            SeedRecognizedHeaders(grid, "Sheet1", binding);
            grid.SetCell("Sheet1", 6, 1, "row-1");
            grid.SetCell("Sheet1", 7, 1, "row-2");
            grid.SetCell("Sheet1", 6, 3, "旧开始时间1");
            grid.SetCell("Sheet1", 6, 4, "旧结束时间1");
            grid.SetCell("Sheet1", 7, 3, "旧开始时间2");
            grid.SetCell("Sheet1", 7, 4, "旧结束时间2");

            var plan = InvokePrepare(service, "PreparePartialDownload", "Sheet1");

            Assert.Equal(0, grid.BeginBulkOperationCount);
            Assert.Equal(0, grid.EndBulkOperationCount);

            InvokeExecute(service, "ExecuteDownload", plan);

            var write = Assert.Single(grid.WriteRangeCalls);
            Assert.Equal("Sheet1", write.SheetName);
            Assert.Equal(6, write.StartRow);
            Assert.Equal(3, write.StartColumn);
            Assert.Equal(2, write.Values.GetLength(0));
            Assert.Equal(2, write.Values.GetLength(1));
            Assert.Equal("2026-02-01", Convert.ToString(write.Values[0, 0]));
            Assert.Equal("2026-02-09", Convert.ToString(write.Values[0, 1]));
            Assert.Equal("2026-03-01", Convert.ToString(write.Values[1, 0]));
            Assert.Equal("2026-03-09", Convert.ToString(write.Values[1, 1]));
            Assert.True(write.WasInsideBulkOperation);
            Assert.Equal(1, grid.BeginBulkOperationCount);
            Assert.Equal(1, grid.EndBulkOperationCount);
            Assert.Equal("2026-02-01", grid.GetCell("Sheet1", 6, 3));
            Assert.Equal("2026-02-09", grid.GetCell("Sheet1", 6, 4));
            Assert.Equal("2026-03-01", grid.GetCell("Sheet1", 7, 3));
            Assert.Equal("2026-03-09", grid.GetCell("Sheet1", 7, 4));
        }

        [Fact]
        public void ExecutePartialDownloadUsesSelectionAreaAndBatchReadsRowIdsForWholeColumnSelection()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
                DataStartRow = 6,
            };
            metadataStore.Bindings["Sheet1"] = binding;
            metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");
            connector.FindResult = new[]
            {
                CreateRow("row-1", "张三", "2026-02-01", "2026-02-09"),
                CreateRow("row-2", "李四", "2026-03-01", "2026-03-09"),
            };

            var selectionReader = new FakeWorksheetSelectionReader
            {
                ThrowOnReadVisibleSelection = true,
                SelectionSnapshot = new WorksheetSelectionSnapshot
                {
                    Areas = new[]
                    {
                        new WorksheetSelectionArea
                        {
                            StartRow = 1,
                            EndRow = 1048576,
                            StartColumn = 3,
                            EndColumn = 4,
                        },
                    },
                },
            };
            var (service, grid) = CreateService(connector, metadataStore, selectionReader);
            SeedRecognizedHeaders(grid, "Sheet1", binding);
            grid.SetCell("Sheet1", 6, 1, "row-1");
            grid.SetCell("Sheet1", 7, 1, "row-2");
            grid.SetCell("Sheet1", 8, 1, string.Empty);
            grid.SetCell("Sheet1", 8, 4, "formatted empty data row");
            grid.SetCell("Sheet1", 6, 3, "旧开始时间1");
            grid.SetCell("Sheet1", 6, 4, "旧结束时间1");
            grid.SetCell("Sheet1", 7, 3, "旧开始时间2");
            grid.SetCell("Sheet1", 7, 4, "旧结束时间2");

            var plan = InvokePrepare(service, "PreparePartialDownload", "Sheet1");

            Assert.Equal("performance", connector.LastFindProjectId);
            Assert.Equal(new[] { "row-1", "row-2" }, connector.LastFindRowIds);
            Assert.Equal(new[] { "start_12345678", "end_12345678" }, connector.LastFindFieldKeys);
            Assert.Contains(grid.ReadRangeCalls, call =>
                call.MethodName == "ReadRangeValues" &&
                call.StartRow == 6 &&
                call.EndRow == 8 &&
                call.StartColumn == 1 &&
                call.EndColumn == 1);
            Assert.DoesNotContain(grid.GetCellTextCalls, call => call.Column == 1 && call.Row >= 6);

            InvokeExecute(service, "ExecuteDownload", plan);

            var write = Assert.Single(grid.WriteRangeCalls);
            Assert.Equal(6, write.StartRow);
            Assert.Equal(3, write.StartColumn);
            Assert.Equal(2, write.Values.GetLength(0));
            Assert.Equal(2, write.Values.GetLength(1));
            Assert.Equal("2026-02-01", grid.GetCell("Sheet1", 6, 3));
            Assert.Equal("2026-02-09", grid.GetCell("Sheet1", 6, 4));
            Assert.Equal("2026-03-01", grid.GetCell("Sheet1", 7, 3));
            Assert.Equal("2026-03-09", grid.GetCell("Sheet1", 7, 4));
            Assert.Equal("formatted empty data row", grid.GetCell("Sheet1", 8, 4));
        }

        [Fact]
        public void ExecutePartialDownloadTreatsWholeSheetSelectionAsAllManagedNonIdFields()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
                DataStartRow = 6,
            };
            metadataStore.Bindings["Sheet1"] = binding;
            metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");
            connector.FindResult = new[] { CreateRow("row-1", "张三", "2026-02-01", "2026-02-09") };

            var selectionReader = new FakeWorksheetSelectionReader
            {
                ThrowOnReadVisibleSelection = true,
                SelectionSnapshot = new WorksheetSelectionSnapshot
                {
                    Areas = new[]
                    {
                        new WorksheetSelectionArea
                        {
                            StartRow = 1,
                            EndRow = 1048576,
                            StartColumn = 1,
                            EndColumn = 16384,
                        },
                    },
                },
            };
            var (service, grid) = CreateService(connector, metadataStore, selectionReader);
            SeedRecognizedHeaders(grid, "Sheet1", binding);
            grid.SetCell("Sheet1", 6, 1, "row-1");
            grid.SetCell("Sheet1", 6, 2, "旧负责人");
            grid.SetCell("Sheet1", 6, 3, "旧开始时间");
            grid.SetCell("Sheet1", 6, 4, "旧结束时间");

            var plan = InvokePrepare(service, "PreparePartialDownload", "Sheet1");

            Assert.Equal(new[] { "row-1" }, connector.LastFindRowIds);
            Assert.Equal(new[] { "owner_name", "start_12345678", "end_12345678" }, connector.LastFindFieldKeys);

            InvokeExecute(service, "ExecuteDownload", plan);

            Assert.Equal("row-1", grid.GetCell("Sheet1", 6, 1));
            Assert.DoesNotContain(grid.WriteRangeCalls, call => call.StartColumn == 1);
            Assert.Equal("张三", grid.GetCell("Sheet1", 6, 2));
            Assert.Equal("2026-02-01", grid.GetCell("Sheet1", 6, 3));
            Assert.Equal("2026-02-09", grid.GetCell("Sheet1", 6, 4));
        }

        [Fact]
        public void ExecutePartialDownloadDoesNotWriteRowsBetweenNonContiguousSelectionAreas()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
                DataStartRow = 6,
            };
            metadataStore.Bindings["Sheet1"] = binding;
            metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");
            connector.FindResult = new[]
            {
                CreateRow("row-1", "张三", "2026-02-01", "2026-02-09"),
                CreateRow("row-2", "李四", "2026-03-01", "2026-03-09"),
                CreateRow("row-3", "王五", "2026-04-01", "2026-04-09"),
            };

            var selectionReader = new FakeWorksheetSelectionReader
            {
                ThrowOnReadVisibleSelection = true,
                SelectionSnapshot = new WorksheetSelectionSnapshot
                {
                    Areas = new[]
                    {
                        new WorksheetSelectionArea { StartRow = 6, EndRow = 6, StartColumn = 3, EndColumn = 3 },
                        new WorksheetSelectionArea { StartRow = 8, EndRow = 8, StartColumn = 3, EndColumn = 3 },
                    },
                },
            };
            var (service, grid) = CreateService(connector, metadataStore, selectionReader);
            SeedRecognizedHeaders(grid, "Sheet1", binding);
            grid.SetCell("Sheet1", 6, 1, "row-1");
            grid.SetCell("Sheet1", 7, 1, "row-2");
            grid.SetCell("Sheet1", 8, 1, "row-3");
            grid.SetCell("Sheet1", 6, 3, "旧开始时间1");
            grid.SetCell("Sheet1", 7, 3, "未选择行");
            grid.SetCell("Sheet1", 8, 3, "旧开始时间3");

            var plan = InvokePrepare(service, "PreparePartialDownload", "Sheet1");

            Assert.Equal(new[] { "row-1", "row-3" }, connector.LastFindRowIds);
            Assert.Equal(new[] { "start_12345678" }, connector.LastFindFieldKeys);

            InvokeExecute(service, "ExecuteDownload", plan);

            Assert.Equal("2026-02-01", grid.GetCell("Sheet1", 6, 3));
            Assert.Equal("未选择行", grid.GetCell("Sheet1", 7, 3));
            Assert.Equal("2026-04-01", grid.GetCell("Sheet1", 8, 3));
        }

        [Fact]
        public void ExecutePartialDownloadDoesNotWriteCrossProductCellsBetweenNonContiguousAreas()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
                DataStartRow = 6,
            };
            metadataStore.Bindings["Sheet1"] = binding;
            metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");
            connector.FindResult = new[]
            {
                CreateRow("row-1", "张三", "2026-02-01", "2026-02-09"),
                CreateRow("row-3", "王五", "2026-04-01", "2026-04-09"),
            };

            var selectionReader = new FakeWorksheetSelectionReader
            {
                ThrowOnReadVisibleSelection = true,
                SelectionSnapshot = new WorksheetSelectionSnapshot
                {
                    Areas = new[]
                    {
                        new WorksheetSelectionArea { StartRow = 6, EndRow = 6, StartColumn = 3, EndColumn = 3 },
                        new WorksheetSelectionArea { StartRow = 8, EndRow = 8, StartColumn = 4, EndColumn = 4 },
                    },
                },
            };
            var (service, grid) = CreateService(connector, metadataStore, selectionReader);
            SeedRecognizedHeaders(grid, "Sheet1", binding);
            grid.SetCell("Sheet1", 6, 1, "row-1");
            grid.SetCell("Sheet1", 8, 1, "row-3");
            grid.SetCell("Sheet1", 6, 3, "旧开始时间1");
            grid.SetCell("Sheet1", 6, 4, "未选择结束时间1");
            grid.SetCell("Sheet1", 8, 3, "未选择开始时间3");
            grid.SetCell("Sheet1", 8, 4, "旧结束时间3");

            var plan = InvokePrepare(service, "PreparePartialDownload", "Sheet1");

            Assert.Equal(new[] { "row-1", "row-3" }, connector.LastFindRowIds);
            Assert.Equal(new[] { "start_12345678", "end_12345678" }, connector.LastFindFieldKeys);

            InvokeExecute(service, "ExecuteDownload", plan);

            Assert.Equal("2026-02-01", grid.GetCell("Sheet1", 6, 3));
            Assert.Equal("未选择结束时间1", grid.GetCell("Sheet1", 6, 4));
            Assert.Equal("未选择开始时间3", grid.GetCell("Sheet1", 8, 3));
            Assert.Equal("2026-04-09", grid.GetCell("Sheet1", 8, 4));
        }

        [Fact]
        public void ExecuteFullDownloadDoesNotRewriteExistingRecognizedHeaders()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
                DataStartRow = 6,
            };
            metadataStore.Bindings["Sheet1"] = binding;
            metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");
            connector.FindResult = new[] { CreateRow("row-1", "张三", "2026-01-02", "2026-01-05") };

            var (service, grid) = CreateService(connector, metadataStore, new FakeWorksheetSelectionReader());
            SeedRecognizedHeaders(grid, "Sheet1", binding);
            grid.SetCell("Sheet1", 6, 1, "row-1");
            grid.SetCell("Sheet1", 6, 2, "旧负责人");
            grid.SetCell("Sheet1", 6, 3, "旧开始");
            grid.SetCell("Sheet1", 6, 4, "旧结束");

            var plan = InvokePrepare(service, "PrepareFullDownload", "Sheet1");
            InvokeExecute(service, "ExecuteDownload", plan);

            Assert.DoesNotContain(grid.ClearedRanges, range => range.StartRow <= 4 && range.EndRow >= 3);
            Assert.Empty(grid.Merges);
            Assert.Equal("ID", grid.GetCell("Sheet1", 3, 1));
            Assert.Equal("项目负责人", grid.GetCell("Sheet1", 3, 2));
            Assert.Equal("测试活动111", grid.GetCell("Sheet1", 3, 3));
            Assert.Equal("开始时间", grid.GetCell("Sheet1", 4, 3));
            Assert.Equal("2026-01-02", grid.GetCell("Sheet1", 6, 3));
        }

        [Fact]
        public void PrepareFullDownloadUsesExistingLayoutWhenGroupedSingleHeadersAreRecognized()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
                DataStartRow = 6,
            };
            metadataStore.Bindings["Sheet1"] = binding;
            metadataStore.FieldMappings["Sheet1"] = BuildGroupedSingleOwnerMappings("Sheet1");
            connector.FindResult = new[] { CreateRow("row-1", "张三", "2026-01-02", "2026-01-05") };

            var (service, grid) = CreateService(connector, metadataStore, new FakeWorksheetSelectionReader());
            SeedGroupedSingleRecognizedHeaders(grid, "Sheet1", binding);
            grid.SetCell("Sheet1", 6, 1, "row-1");

            var plan = InvokePrepare(service, "PrepareFullDownload", "Sheet1");

            Assert.True(ReadBoolProperty(plan, "UsesExistingLayout"));
        }

        [Fact]
        public void ExecuteFullDownloadWithEmptyHeadersFlattensGroupedSingleToChildText()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
                DataStartRow = 6,
            };
            metadataStore.Bindings["Sheet1"] = binding;
            metadataStore.FieldMappings["Sheet1"] = BuildGroupedSingleOwnerMappings("Sheet1");
            connector.FindResult = new[] { CreateRow("row-1", "张三", "2026-01-02", "2026-01-05") };

            var (service, grid) = CreateService(connector, metadataStore, new FakeWorksheetSelectionReader());

            var plan = InvokePrepare(service, "PrepareFullDownload", "Sheet1");
            InvokeExecute(service, "ExecuteDownload", plan);

            Assert.Equal("ID", grid.GetCell("Sheet1", 3, 1));
            Assert.Equal("负责人", grid.GetCell("Sheet1", 3, 2));
            Assert.NotEqual("联系人信息", grid.GetCell("Sheet1", 3, 2));
            Assert.Equal("测试活动111", grid.GetCell("Sheet1", 3, 3));
            Assert.Equal("开始时间", grid.GetCell("Sheet1", 4, 3));
        }

        [Fact]
        public void ExecuteFullDownloadUsesBatchWriteForContiguousManagedColumns()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
                DataStartRow = 6,
            };
            metadataStore.Bindings["Sheet1"] = binding;
            metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");
            connector.FindResult = new[]
            {
                CreateRow("row-1", "张三", "2026-01-02", "2026-01-05"),
                CreateRow("row-2", "李四", "2026-02-03", "2026-02-06"),
            };

            var (service, grid) = CreateService(connector, metadataStore, new FakeWorksheetSelectionReader());
            SeedRecognizedHeaders(grid, "Sheet1", binding);

            var plan = InvokePrepare(service, "PrepareFullDownload", "Sheet1");
            InvokeExecute(service, "ExecuteDownload", plan);

            var write = Assert.Single(grid.WriteRangeCalls);
            Assert.Equal("Sheet1", write.SheetName);
            Assert.Equal(6, write.StartRow);
            Assert.Equal(1, write.StartColumn);
            Assert.Equal(2, write.Values.GetLength(0));
            Assert.Equal(4, write.Values.GetLength(1));
            Assert.Equal("row-1", Convert.ToString(write.Values[0, 0]));
            Assert.Equal("张三", Convert.ToString(write.Values[0, 1]));
            Assert.Equal("2026-02-03", Convert.ToString(write.Values[1, 2]));
            Assert.Equal("2026-02-06", Convert.ToString(write.Values[1, 3]));
            Assert.Equal(1, grid.BeginBulkOperationCount);
            Assert.Equal(1, grid.EndBulkOperationCount);
        }

        [Fact]
        public void ExecuteFullDownloadBeginsAndEndsOneBulkOperation()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
                DataStartRow = 6,
            };
            metadataStore.Bindings["Sheet1"] = binding;
            metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");
            connector.FindResult = new[]
            {
                CreateRow("row-1", "张三", "2026-01-02", "2026-01-05"),
            };

            var (service, grid) = CreateService(connector, metadataStore, new FakeWorksheetSelectionReader());
            SeedRecognizedHeaders(grid, "Sheet1", binding);

            var plan = InvokePrepare(service, "PrepareFullDownload", "Sheet1");

            Assert.Equal(0, grid.BeginBulkOperationCount);
            Assert.Equal(0, grid.EndBulkOperationCount);

            InvokeExecute(service, "ExecuteDownload", plan);

            Assert.Equal(1, grid.BeginBulkOperationCount);
            Assert.Equal(1, grid.EndBulkOperationCount);
            Assert.Contains(grid.LastUsedRowCalls, call => call.SheetName == "Sheet1" && call.WasInsideBulkOperation);
            Assert.Contains(grid.WriteRangeCalls, call => call.SheetName == "Sheet1" && call.WasInsideBulkOperation);
        }

        [Fact]
        public void ExecuteFullDownloadSplitsBatchWritesAcrossNonContiguousManagedColumns()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
                DataStartRow = 6,
            };
            metadataStore.Bindings["Sheet1"] = binding;
            metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");
            connector.FindResult = new[] { CreateRow("row-1", "张三", "2026-01-02", "2026-01-05") };

            var (service, grid) = CreateService(connector, metadataStore, new FakeWorksheetSelectionReader());
            grid.SetCell("Sheet1", 3, 1, "ID");
            grid.SetCell("Sheet1", 3, 2, "项目负责人");
            grid.SetCell("Sheet1", 3, 3, "用户备注");
            grid.SetCell("Sheet1", 3, 4, "测试活动111");
            grid.SetCell("Sheet1", 4, 4, "开始时间");
            grid.SetCell("Sheet1", 4, 5, "结束时间");
            grid.SetCell("Sheet1", 6, 3, "保留的备注");

            var plan = InvokePrepare(service, "PrepareFullDownload", "Sheet1");
            InvokeExecute(service, "ExecuteDownload", plan);

            Assert.Equal(2, grid.WriteRangeCalls.Count);
            Assert.Equal(1, grid.WriteRangeCalls[0].StartColumn);
            Assert.Equal(2, grid.WriteRangeCalls[0].Values.GetLength(1));
            Assert.Equal("row-1", Convert.ToString(grid.WriteRangeCalls[0].Values[0, 0]));
            Assert.Equal("张三", Convert.ToString(grid.WriteRangeCalls[0].Values[0, 1]));
            Assert.Equal(4, grid.WriteRangeCalls[1].StartColumn);
            Assert.Equal(2, grid.WriteRangeCalls[1].Values.GetLength(1));
            Assert.Equal("2026-01-02", Convert.ToString(grid.WriteRangeCalls[1].Values[0, 0]));
            Assert.Equal("2026-01-05", Convert.ToString(grid.WriteRangeCalls[1].Values[0, 1]));
            Assert.Equal("保留的备注", grid.GetCell("Sheet1", 6, 3));
            Assert.Equal(1, grid.BeginBulkOperationCount);
            Assert.Equal(1, grid.EndBulkOperationCount);
        }

        [Fact]
        public void ExecuteFullUploadUsesConfiguredDataStartRowAndRecognizedColumns()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
                DataStartRow = 6,
            };
            metadataStore.Bindings["Sheet1"] = binding;
            metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");

            var (service, grid) = CreateService(connector, metadataStore, new FakeWorksheetSelectionReader());
            SeedRecognizedHeaders(grid, "Sheet1", binding);
            grid.SetCell("Sheet1", 5, 1, "统计行");
            grid.SetCell("Sheet1", 6, 1, "row-1");
            grid.SetCell("Sheet1", 6, 2, "李四");
            grid.SetCell("Sheet1", 6, 3, "2026-01-02");
            grid.SetCell("Sheet1", 6, 4, "2026-01-05");
            grid.SetCell("Sheet1", 7, 1, string.Empty);
            grid.SetCell("Sheet1", 7, 2, "无ID");
            grid.SetCell("Sheet1", 7, 3, "2026-03-01");
            grid.SetCell("Sheet1", 7, 4, "2026-03-05");

            var plan = InvokePrepare(service, "PrepareFullUpload", "Sheet1");
            var preview = ReadPreview(plan);
            Assert.Equal(3, preview.Changes.Length);

            InvokeExecute(service, "ExecuteUpload", plan);

            Assert.Equal("performance", connector.LastBatchSaveProjectId);
            Assert.Equal(3, connector.LastBatchSaveChanges.Count);
            Assert.Contains(connector.LastBatchSaveChanges, change => change.RowId == "row-1" && change.ApiFieldKey == "owner_name" && change.NewValue == "李四");
            Assert.Contains(connector.LastBatchSaveChanges, change => change.RowId == "row-1" && change.ApiFieldKey == "start_12345678" && change.NewValue == "2026-01-02");
            Assert.Contains(connector.LastBatchSaveChanges, change => change.RowId == "row-1" && change.ApiFieldKey == "end_12345678" && change.NewValue == "2026-01-05");
            Assert.DoesNotContain(connector.LastBatchSaveChanges, change => string.IsNullOrWhiteSpace(change.RowId));
        }

        [Fact]
        public void PrepareFullUploadBeginsAndEndsOneBulkOperation()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
                DataStartRow = 6,
            };
            metadataStore.Bindings["Sheet1"] = binding;
            metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");

            var (service, grid) = CreateService(connector, metadataStore, new FakeWorksheetSelectionReader());
            SeedRecognizedHeaders(grid, "Sheet1", binding);
            grid.SetCell("Sheet1", 6, 1, "row-1");
            grid.SetCell("Sheet1", 6, 2, "李四");
            grid.SetCell("Sheet1", 6, 3, "2026-01-02");
            grid.SetCell("Sheet1", 6, 4, "2026-01-05");

            Assert.Equal(0, grid.BeginBulkOperationCount);
            Assert.Equal(0, grid.EndBulkOperationCount);

            var plan = InvokePrepare(service, "PrepareFullUpload", "Sheet1");

            Assert.NotNull(plan);
            Assert.Equal(1, grid.BeginBulkOperationCount);
            Assert.Equal(1, grid.EndBulkOperationCount);
            Assert.Contains(grid.ReadRangeCalls, call => call.MethodName == "ReadRangeValues" && call.WasInsideBulkOperation);
            Assert.Contains(grid.ReadRangeCalls, call => call.MethodName == "ReadRangeNumberFormats" && call.WasInsideBulkOperation);
        }

        [Fact]
        public void ExecuteFullUploadUsesBatchReadForManagedRegion()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
                DataStartRow = 6,
            };
            metadataStore.Bindings["Sheet1"] = binding;
            metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");

            var (service, grid) = CreateService(connector, metadataStore, new FakeWorksheetSelectionReader());
            SeedRecognizedHeaders(grid, "Sheet1", binding);
            grid.SetRawCell("Sheet1", 6, 1, "row-1");
            grid.SetRawCell("Sheet1", 6, 2, "李四");
            grid.SetRawCell("Sheet1", 6, 3, 1234d, "General", "001234");
            grid.SetRawCell("Sheet1", 6, 4, 56.75d, "General", "56.75-显示");
            grid.SetRawCell("Sheet1", 7, 1, string.Empty);
            grid.SetRawCell("Sheet1", 7, 2, "无ID");
            grid.SetRawCell("Sheet1", 7, 3, 999d, "General", "999");
            grid.SetRawCell("Sheet1", 7, 4, 1000d, "General", "1000");

            var plan = InvokePrepare(service, "PrepareFullUpload", "Sheet1");
            InvokeExecute(service, "ExecuteUpload", plan);

            Assert.Collection(
                grid.ReadRangeCalls,
                call =>
                {
                    Assert.Equal("ReadRangeValues", call.MethodName);
                    Assert.Equal("Sheet1", call.SheetName);
                    Assert.Equal(6, call.StartRow);
                    Assert.Equal(7, call.EndRow);
                    Assert.Equal(1, call.StartColumn);
                    Assert.Equal(4, call.EndColumn);
                },
                call =>
                {
                    Assert.Equal("ReadRangeNumberFormats", call.MethodName);
                    Assert.Equal("Sheet1", call.SheetName);
                    Assert.Equal(6, call.StartRow);
                    Assert.Equal(7, call.EndRow);
                    Assert.Equal(1, call.StartColumn);
                    Assert.Equal(4, call.EndColumn);
                });
            Assert.Equal(0, grid.CountGetCellTextCalls("Sheet1", 6, 1));
            Assert.Equal(0, grid.CountGetCellTextCalls("Sheet1", 6, 2));
            Assert.Equal(0, grid.CountGetCellTextCalls("Sheet1", 6, 3));
            Assert.Equal(0, grid.CountGetCellTextCalls("Sheet1", 6, 4));
            Assert.Contains(connector.LastBatchSaveChanges, change => change.RowId == "row-1" && change.ApiFieldKey == "start_12345678" && change.NewValue == "1234");
            Assert.Contains(connector.LastBatchSaveChanges, change => change.RowId == "row-1" && change.ApiFieldKey == "end_12345678" && change.NewValue == "56.75");
        }

        [Fact]
        public void ExecuteFullUploadSplitsBatchReadsAcrossNonContiguousManagedColumns()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
                DataStartRow = 6,
            };
            metadataStore.Bindings["Sheet1"] = binding;
            metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");

            var (service, grid) = CreateService(connector, metadataStore, new FakeWorksheetSelectionReader());
            grid.SetCell("Sheet1", 3, 1, "ID");
            grid.SetCell("Sheet1", 3, 2, "项目负责人");
            grid.SetCell("Sheet1", 3, 3, "用户备注");
            grid.SetCell("Sheet1", 3, 4, "测试活动111");
            grid.SetCell("Sheet1", 4, 4, "开始时间");
            grid.SetCell("Sheet1", 4, 5, "结束时间");
            grid.SetRawCell("Sheet1", 6, 1, "row-1");
            grid.SetRawCell("Sheet1", 6, 2, "李四");
            grid.SetRawCell("Sheet1", 6, 3, "保留备注");
            grid.SetRawCell("Sheet1", 6, 4, 1234d, "General", "001234");
            grid.SetRawCell("Sheet1", 6, 5, 56.75d, "General", "56.75-显示");

            var plan = InvokePrepare(service, "PrepareFullUpload", "Sheet1");
            InvokeExecute(service, "ExecuteUpload", plan);

            Assert.Collection(
                grid.ReadRangeCalls,
                call =>
                {
                    Assert.Equal("ReadRangeValues", call.MethodName);
                    Assert.Equal("Sheet1", call.SheetName);
                    Assert.Equal(6, call.StartRow);
                    Assert.Equal(6, call.EndRow);
                    Assert.Equal(1, call.StartColumn);
                    Assert.Equal(2, call.EndColumn);
                },
                call =>
                {
                    Assert.Equal("ReadRangeNumberFormats", call.MethodName);
                    Assert.Equal("Sheet1", call.SheetName);
                    Assert.Equal(6, call.StartRow);
                    Assert.Equal(6, call.EndRow);
                    Assert.Equal(1, call.StartColumn);
                    Assert.Equal(2, call.EndColumn);
                },
                call =>
                {
                    Assert.Equal("ReadRangeValues", call.MethodName);
                    Assert.Equal("Sheet1", call.SheetName);
                    Assert.Equal(6, call.StartRow);
                    Assert.Equal(6, call.EndRow);
                    Assert.Equal(4, call.StartColumn);
                    Assert.Equal(5, call.EndColumn);
                },
                call =>
                {
                    Assert.Equal("ReadRangeNumberFormats", call.MethodName);
                    Assert.Equal("Sheet1", call.SheetName);
                    Assert.Equal(6, call.StartRow);
                    Assert.Equal(6, call.EndRow);
                    Assert.Equal(4, call.StartColumn);
                    Assert.Equal(5, call.EndColumn);
                });
            Assert.DoesNotContain(grid.ReadRangeCalls, call => call.StartColumn <= 3 && call.EndColumn >= 3);
            Assert.Equal(0, grid.CountGetCellTextCalls("Sheet1", 6, 3));
            Assert.Contains(connector.LastBatchSaveChanges, change => change.RowId == "row-1" && change.ApiFieldKey == "owner_name" && change.NewValue == "李四");
            Assert.Contains(connector.LastBatchSaveChanges, change => change.RowId == "row-1" && change.ApiFieldKey == "start_12345678" && change.NewValue == "1234");
            Assert.Contains(connector.LastBatchSaveChanges, change => change.RowId == "row-1" && change.ApiFieldKey == "end_12345678" && change.NewValue == "56.75");
        }

        [Fact]
        public void ExecuteFullUploadFallsBackToCellTextForUnsafeFormattedCells()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
                DataStartRow = 6,
            };
            metadataStore.Bindings["Sheet1"] = binding;
            metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");

            var (service, grid) = CreateService(connector, metadataStore, new FakeWorksheetSelectionReader());
            SeedRecognizedHeaders(grid, "Sheet1", binding);
            grid.SetRawCell("Sheet1", 6, 1, "row-1");
            grid.SetRawCell("Sheet1", 6, 2, "李四");
            grid.SetRawCell("Sheet1", 6, 3, 45734d, "yyyy-mm-dd", "2025-03-18");
            grid.SetRawCell("Sheet1", 6, 4, 0.25d, "0%", "25%");

            var plan = InvokePrepare(service, "PrepareFullUpload", "Sheet1");
            InvokeExecute(service, "ExecuteUpload", plan);

            Assert.Collection(
                grid.ReadRangeCalls,
                call => Assert.Equal("ReadRangeValues", call.MethodName),
                call => Assert.Equal("ReadRangeNumberFormats", call.MethodName));
            Assert.Equal(0, grid.CountGetCellTextCalls("Sheet1", 6, 1));
            Assert.Equal(0, grid.CountGetCellTextCalls("Sheet1", 6, 2));
            Assert.Equal(1, grid.CountGetCellTextCalls("Sheet1", 6, 3));
            Assert.Equal(1, grid.CountGetCellTextCalls("Sheet1", 6, 4));
            Assert.Contains(connector.LastBatchSaveChanges, change => change.RowId == "row-1" && change.ApiFieldKey == "start_12345678" && change.NewValue == "2025-03-18");
            Assert.Contains(connector.LastBatchSaveChanges, change => change.RowId == "row-1" && change.ApiFieldKey == "end_12345678" && change.NewValue == "25%");
        }

        [Fact]
        public void PrepareFullDownloadRequiresExplicitInitializationWhenStoredMappingsAreUnusable()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
                DataStartRow = 6,
            };
            metadataStore.Bindings["Sheet1"] = binding;
            metadataStore.FieldMappings["Sheet1"] = BuildLegacyMappingsWithoutIdFlag("Sheet1");
            connector.FindResult = new[] { CreateRow("row-1", "张三", "2026-01-02", "2026-01-05") };

            var (service, _) = CreateService(connector, metadataStore, new FakeWorksheetSelectionReader());

            var exception = Assert.Throws<TargetInvocationException>(() => InvokePrepare(service, "PrepareFullDownload", "Sheet1"));
            var inner = Assert.IsType<InvalidOperationException>(exception.InnerException);
            Assert.Contains("The current sheet is not initialized. Initialize the current sheet first.", inner.Message);
            Assert.Null(metadataStore.LastSavedBinding);
            Assert.Empty(metadataStore.LastSavedFieldMappings);
        }

        [Fact]
        public void PrepareFullDownloadRequiresExplicitInitializationWhenFieldMappingsAreMissing()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            metadataStore.Bindings["Sheet1"] = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "new-project",
                ProjectName = "新项目",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
                DataStartRow = 6,
            };

            var (service, _) = CreateService(connector, metadataStore, new FakeWorksheetSelectionReader());

            var exception = Assert.Throws<TargetInvocationException>(() => InvokePrepare(service, "PrepareFullDownload", "Sheet1"));
            var inner = Assert.IsType<InvalidOperationException>(exception.InnerException);
            Assert.Contains("The current sheet is not initialized. Initialize the current sheet first.", inner.Message);
        }

        [Fact]
        public void ExecutePartialUploadUsesRecognizedHeadersAndIdLookupOutsideSelection()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
                DataStartRow = 6,
            };
            metadataStore.Bindings["Sheet1"] = binding;
            metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");

            var selectionReader = new FakeWorksheetSelectionReader
            {
                VisibleCells = new[]
                {
                    new SelectedVisibleCell { Row = 6, Column = 4, Value = "2026-01-10" },
                },
            };
            var (service, grid) = CreateService(connector, metadataStore, selectionReader);
            SeedRecognizedHeaders(grid, "Sheet1", binding);
            grid.SetCell("Sheet1", 6, 1, "row-1");
            grid.SetCell("Sheet1", 6, 2, "张三");
            grid.SetCell("Sheet1", 6, 3, "2026-01-02");
            grid.SetCell("Sheet1", 6, 4, "2026-01-10");

            var plan = InvokePrepare(service, "PreparePartialUpload", "Sheet1");
            var preview = ReadPreview(plan);
            Assert.Single(preview.Changes);
            Assert.Equal("row-1", preview.Changes[0].RowId);
            Assert.Equal("end_12345678", preview.Changes[0].ApiFieldKey);

            InvokeExecute(service, "ExecuteUpload", plan);

            Assert.Equal("performance", connector.LastBatchSaveProjectId);
            Assert.Single(connector.LastBatchSaveChanges);
            Assert.Equal("2026-01-10", connector.LastBatchSaveChanges[0].NewValue);
        }

        [Fact]
        public void PreparePartialUploadFiltersPreviewAndExecuteUploadSubmitsOnlyIncludedChanges()
        {
            var connector = new FakeSystemConnector
            {
                SkippedApiFieldKey = "end_12345678",
                SkipReason = "单据已归档，禁止上传",
            };
            var metadataStore = new FakeWorksheetMetadataStore();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
                DataStartRow = 6,
            };
            metadataStore.Bindings["Sheet1"] = binding;
            metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");

            var selectionReader = new FakeWorksheetSelectionReader
            {
                VisibleCells = new[]
                {
                    new SelectedVisibleCell { Row = 6, Column = 2, Value = "李四" },
                    new SelectedVisibleCell { Row = 6, Column = 4, Value = "2026-01-10" },
                },
            };
            var (service, grid) = CreateService(connector, metadataStore, selectionReader);
            SeedRecognizedHeaders(grid, "Sheet1", binding);
            grid.SetCell("Sheet1", 6, 1, "row-1");
            grid.SetCell("Sheet1", 6, 2, "李四");
            grid.SetCell("Sheet1", 6, 4, "2026-01-10");

            var plan = InvokePrepare(service, "PreparePartialUpload", "Sheet1");
            var preview = ReadPreview(plan);

            Assert.Equal("performance", connector.LastFilterProjectId);
            Assert.Equal("Upload will submit 1 cell(s) and skip 1 cell(s).", preview.Summary);
            var included = Assert.Single(preview.Changes);
            Assert.Equal("owner_name", included.ApiFieldKey);
            var skipped = Assert.Single(preview.SkippedChanges);
            Assert.Equal("end_12345678", skipped.Change.ApiFieldKey);
            Assert.Equal("单据已归档，禁止上传", skipped.Reason);
            Assert.Contains("row-1 / end_12345678: Skipped, 单据已归档，禁止上传", preview.Details);

            InvokeExecute(service, "ExecuteUpload", plan);

            Assert.Single(connector.LastBatchSaveChanges);
            Assert.Equal("owner_name", connector.LastBatchSaveChanges[0].ApiFieldKey);
        }

        [Fact]
        public void ExecutePartialUploadAppendsWorkbookLogAfterSuccessfulBatchSave()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
                DataStartRow = 6,
            };
            metadataStore.Bindings["Sheet1"] = binding;
            metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");

            var selectionReader = new FakeWorksheetSelectionReader
            {
                VisibleCells = new[]
                {
                    new SelectedVisibleCell { Row = 6, Column = 4, Value = "2026-01-10" },
                },
            };
            var (service, grid, logStore, pendingEditTracker) = CreateServiceWithChangeLog(connector, metadataStore, selectionReader);
            SeedRecognizedHeaders(grid, "Sheet1", binding);
            grid.SetCell("Sheet1", 6, 1, "row-1");
            grid.SetCell("Sheet1", 6, 4, "2026-01-10");
            SeedPendingOriginalValue(pendingEditTracker, "Sheet1", 6, 4, "2026-01-05");

            var plan = InvokePrepare(service, "PreparePartialUpload", "Sheet1");
            InvokeExecute(service, "ExecuteUpload", plan);

            var entry = Assert.Single(logStore.Entries);
            Assert.Equal("row-1", entry.Key);
            Assert.Equal("测试活动111/结束时间", entry.HeaderText);
            Assert.Equal("Upload", entry.ChangeMode);
            Assert.Equal("2026-01-10", entry.NewValue);
            Assert.Equal("2026-01-05", entry.OldValue);
            Assert.False(TryGetPendingOriginalValue(pendingEditTracker, "Sheet1", 6, 4, out _));
        }

        [Fact]
        public void ExecutePartialUploadDoesNotAppendLogOrClearPendingValueWhenBatchSaveFails()
        {
            var connector = new FakeSystemConnector
            {
                BatchSaveException = new InvalidOperationException("save failed"),
            };
            var metadataStore = new FakeWorksheetMetadataStore();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
                DataStartRow = 6,
            };
            metadataStore.Bindings["Sheet1"] = binding;
            metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");

            var selectionReader = new FakeWorksheetSelectionReader
            {
                VisibleCells = new[]
                {
                    new SelectedVisibleCell { Row = 6, Column = 4, Value = "2026-01-10" },
                },
            };
            var (service, grid, logStore, pendingEditTracker) = CreateServiceWithChangeLog(connector, metadataStore, selectionReader);
            SeedRecognizedHeaders(grid, "Sheet1", binding);
            grid.SetCell("Sheet1", 6, 1, "row-1");
            grid.SetCell("Sheet1", 6, 4, "2026-01-10");
            SeedPendingOriginalValue(pendingEditTracker, "Sheet1", 6, 4, "2026-01-05");

            var plan = InvokePrepare(service, "PreparePartialUpload", "Sheet1");

            Assert.Throws<TargetInvocationException>(() => InvokeExecute(service, "ExecuteUpload", plan));
            Assert.Empty(logStore.Entries);
            Assert.True(TryGetPendingOriginalValue(pendingEditTracker, "Sheet1", 6, 4, out var value));
            Assert.Equal("2026-01-05", value);
        }

        [Fact]
        public void PreparePartialUploadResolvesGroupedSingleOwnerNameFromTwoRowHeader()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
                DataStartRow = 6,
            };
            metadataStore.Bindings["Sheet1"] = binding;
            metadataStore.FieldMappings["Sheet1"] = BuildGroupedSingleOwnerMappings("Sheet1");

            var selectionReader = new FakeWorksheetSelectionReader
            {
                VisibleCells = new[]
                {
                    new SelectedVisibleCell { Row = 6, Column = 2, Value = "李四" },
                },
            };
            var (service, grid) = CreateService(connector, metadataStore, selectionReader);
            SeedGroupedSingleRecognizedHeaders(grid, "Sheet1", binding);
            grid.SetCell("Sheet1", 6, 1, "row-1");
            grid.SetCell("Sheet1", 6, 2, "李四");

            var plan = InvokePrepare(service, "PreparePartialUpload", "Sheet1");
            var preview = ReadPreview(plan);
            var change = Assert.Single(preview.Changes);
            Assert.Equal("owner_name", change.ApiFieldKey);
            Assert.Equal("row-1", change.RowId);
            Assert.Equal("李四", change.NewValue);
        }

        [Fact]
        public void PreparePartialUploadReadsEachRowIdAtMostOncePerRow()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
                DataStartRow = 6,
            };
            metadataStore.Bindings["Sheet1"] = binding;
            metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");

            var selectionReader = new FakeWorksheetSelectionReader
            {
                VisibleCells = new[]
                {
                    new SelectedVisibleCell { Row = 6, Column = 3, Value = "2026-01-02" },
                    new SelectedVisibleCell { Row = 6, Column = 4, Value = "2026-01-05" },
                    new SelectedVisibleCell { Row = 7, Column = 2, Value = "王五" },
                },
            };
            var (service, grid) = CreateService(connector, metadataStore, selectionReader);
            SeedRecognizedHeaders(grid, "Sheet1", binding);
            grid.SetCell("Sheet1", 6, 1, "row-1");
            grid.SetCell("Sheet1", 6, 2, "张三");
            grid.SetCell("Sheet1", 6, 3, "2026-01-02");
            grid.SetCell("Sheet1", 6, 4, "2026-01-05");
            grid.SetCell("Sheet1", 7, 1, "row-2");
            grid.SetCell("Sheet1", 7, 2, "王五");

            var plan = InvokePrepare(service, "PreparePartialUpload", "Sheet1");
            var preview = ReadPreview(plan);

            Assert.Equal(3, preview.Changes.Length);
            Assert.Equal(1, grid.CountGetCellTextCalls("Sheet1", 6, 1));
            Assert.Equal(1, grid.CountGetCellTextCalls("Sheet1", 7, 1));
        }

        [Fact]
        public void ExecutePartialDownloadReadsEachRowIdAtMostOncePerRow()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
                DataStartRow = 6,
            };
            metadataStore.Bindings["Sheet1"] = binding;
            metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");
            connector.FindResult = new[]
            {
                CreateRow("row-1", "张三", "2026-02-01", "2026-02-09"),
                CreateRow("row-2", "王五", "2026-03-01", "2026-03-07"),
            };

            var selectionReader = new FakeWorksheetSelectionReader
            {
                VisibleCells = new[]
                {
                    new SelectedVisibleCell { Row = 6, Column = 3, Value = "旧开始时间" },
                    new SelectedVisibleCell { Row = 6, Column = 4, Value = "旧结束时间" },
                    new SelectedVisibleCell { Row = 7, Column = 2, Value = "旧负责人" },
                },
            };
            var (service, grid) = CreateService(connector, metadataStore, selectionReader);
            SeedRecognizedHeaders(grid, "Sheet1", binding);
            grid.SetCell("Sheet1", 6, 1, "row-1");
            grid.SetCell("Sheet1", 6, 2, "张三");
            grid.SetCell("Sheet1", 6, 3, "旧开始时间");
            grid.SetCell("Sheet1", 6, 4, "旧结束时间");
            grid.SetCell("Sheet1", 7, 1, "row-2");
            grid.SetCell("Sheet1", 7, 2, "旧负责人");

            var plan = InvokePrepare(service, "PreparePartialDownload", "Sheet1");
            grid.GetCellTextCalls.Clear();

            InvokeExecute(service, "ExecuteDownload", plan);

            Assert.Equal(1, grid.CountGetCellTextCalls("Sheet1", 6, 1));
            Assert.Equal(1, grid.CountGetCellTextCalls("Sheet1", 7, 1));
            Assert.Equal("2026-02-01", grid.GetCell("Sheet1", 6, 3));
            Assert.Equal("2026-02-09", grid.GetCell("Sheet1", 6, 4));
            Assert.Equal("王五", grid.GetCell("Sheet1", 7, 2));
        }

        [Fact]
        public void ExecuteFullDownloadUsesActiveWorkbookMetadataWhenDifferentWorkbooksShareSameSheetName()
        {
            var connector = new FakeSystemConnector();
            var adapter = CreateScopedMetadataAdapter();
            SeedWorkbookMetadata(
                adapter,
                "WorkbookA",
                new SheetBinding
                {
                    SheetName = "Sheet1",
                    SystemKey = "current-business-system",
                    ProjectId = "project-a",
                    ProjectName = "项目A",
                    HeaderStartRow = 3,
                    HeaderRowCount = 2,
                    DataStartRow = 6,
                },
                BuildDefaultMappings("Sheet1"));
            SeedWorkbookMetadata(
                adapter,
                "WorkbookB",
                new SheetBinding
                {
                    SheetName = "Sheet1",
                    SystemKey = "current-business-system",
                    ProjectId = "project-b",
                    ProjectName = "项目B",
                    HeaderStartRow = 5,
                    HeaderRowCount = 2,
                    DataStartRow = 9,
                },
                BuildDefaultMappings("Sheet1"));

            var metadataStore = CreateRealMetadataStore(adapter);
            var (service, grid) = CreateService(connector, metadataStore, new FakeWorksheetSelectionReader());

            adapter.SwitchWorkbook("WorkbookA");
            connector.FindResult = new[] { CreateRow("row-a", "张三", "2026-01-02", "2026-01-05") };

            var workbookAPlan = InvokePrepare(service, "PrepareFullDownload", "Sheet1");
            InvokeExecute(service, "ExecuteDownload", workbookAPlan);

            Assert.Equal("project-a", connector.LastFindProjectId);
            Assert.Equal("row-a", grid.GetCell("Sheet1", 6, 1));

            grid.ClearAllCells();
            adapter.SwitchWorkbook("WorkbookB");
            connector.FindResult = new[] { CreateRow("row-b", "李四", "2026-02-02", "2026-02-06") };

            var workbookBPlan = InvokePrepare(service, "PrepareFullDownload", "Sheet1");
            InvokeExecute(service, "ExecuteDownload", workbookBPlan);

            Assert.Equal("project-b", connector.LastFindProjectId);
            Assert.Equal("ID", grid.GetCell("Sheet1", 5, 1));
            Assert.Equal("项目负责人", grid.GetCell("Sheet1", 5, 2));
            Assert.Equal("row-b", grid.GetCell("Sheet1", 9, 1));
            Assert.Equal("李四", grid.GetCell("Sheet1", 9, 2));
            Assert.Equal(string.Empty, grid.GetCell("Sheet1", 6, 1));
        }

        [Fact]
        public void PrepareAiColumnMappingPreviewScansFullHeaderAreaAndCallsClient()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
                DataStartRow = 6,
            };
            metadataStore.Bindings["Sheet1"] = binding;
            metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");
            var aiClient = new FakeAiColumnMappingClient
            {
                Response = new AiColumnMappingResponse
                {
                    Mappings = new[]
                    {
                        new AiColumnMappingSuggestion
                        {
                            ExcelColumn = 2,
                            ActualL1 = "基础信息",
                            ActualL2 = "负责人",
                            TargetHeaderId = "owner_name",
                            TargetApiFieldKey = "owner_name",
                            Confidence = 0.91,
                        },
                    },
                },
            };

            var (service, grid) = CreateService(connector, metadataStore, new FakeWorksheetSelectionReader(), aiClient);
            grid.SetCell("Sheet1", 3, 1, "ID");
            grid.SetCell("Sheet1", 3, 2, "基础信息");
            grid.SetCell("Sheet1", 4, 2, "负责人");
            grid.SetCell("Sheet1", 3, 3, "测试活动111");
            grid.SetCell("Sheet1", 4, 3, "开始时间");
            grid.SetCell("Sheet1", 4, 4, "结束时间");

            var preview = (AiColumnMappingPreview)InvokePrepare(service, "PrepareAiColumnMappingPreview", "Sheet1");

            Assert.Equal("Sheet1", aiClient.LastRequest.SheetName);
            var actualHeader = Assert.Single(aiClient.LastRequest.ActualHeaders);
            Assert.Equal(2, actualHeader.ExcelColumn);
            Assert.Equal("基础信息", actualHeader.ActualL1);
            Assert.Equal("负责人", actualHeader.ActualL2);
            Assert.DoesNotContain(aiClient.LastRequest.ActualHeaders, header => header.ExcelColumn == 4);
            Assert.DoesNotContain(aiClient.LastRequest.Candidates, candidate => string.Equals(candidate.HeaderId, "other_sheet_owner", StringComparison.Ordinal));
            var item = Assert.Single(preview.Items);
            Assert.Equal(AiColumnMappingPreviewStatuses.Accepted, item.Status);
            Assert.Equal("基础信息", item.ActualL1);
            Assert.Equal("负责人", item.ActualL2);
        }

        [Fact]
        public void PrepareAiColumnMappingPreviewDoesNotCallClientForAlreadyMatchedHeaders()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            metadataStore.Bindings["Sheet1"] = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
                DataStartRow = 6,
            };
            metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");
            var aiClient = new FakeAiColumnMappingClient();

            var (service, grid) = CreateService(connector, metadataStore, new FakeWorksheetSelectionReader(), aiClient);
            grid.SetCell("Sheet1", 3, 1, "ID");
            grid.SetCell("Sheet1", 3, 2, "项目负责人");
            grid.SetCell("Sheet1", 3, 3, "测试活动111");
            grid.SetCell("Sheet1", 4, 3, "开始时间");
            grid.SetCell("Sheet1", 4, 4, "结束时间");

            var preview = (AiColumnMappingPreview)InvokePrepare(service, "PrepareAiColumnMappingPreview", "Sheet1");

            Assert.Null(aiClient.LastRequest);
            Assert.Empty(preview.Items);
        }

        [Fact]
        public async Task PrepareAiColumnMappingPreviewAsyncPassesCancellationTokenToClient()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
                DataStartRow = 6,
            };
            metadataStore.Bindings["Sheet1"] = binding;
            metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1");
            var aiClient = new FakeAiColumnMappingClient
            {
                Response = new AiColumnMappingResponse
                {
                    Mappings = new[]
                    {
                        new AiColumnMappingSuggestion
                        {
                            ExcelColumn = 2,
                            ActualL1 = "基础信息",
                            ActualL2 = "负责人",
                            TargetHeaderId = "owner_name",
                            TargetApiFieldKey = "owner_name",
                            Confidence = 0.91,
                        },
                    },
                },
            };

            var (service, grid) = CreateService(connector, metadataStore, new FakeWorksheetSelectionReader(), aiClient);
            grid.SetCell("Sheet1", 3, 1, "ID");
            grid.SetCell("Sheet1", 3, 2, "基础信息");
            grid.SetCell("Sheet1", 4, 2, "负责人");
            using (var cancellationTokenSource = new CancellationTokenSource())
            {
                var preview = await InvokePrepareAiColumnMappingPreviewAsync(service, "Sheet1", cancellationTokenSource.Token);

                Assert.Equal(cancellationTokenSource.Token, aiClient.LastCancellationToken);
                Assert.Single(preview.Items);
            }
        }

        [Fact]
        public void ApplyAiColumnMappingPreviewSavesOnlyConfirmedMetadataRows()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
                DataStartRow = 6,
            };
            metadataStore.Bindings["Sheet1"] = binding;
            metadataStore.FieldMappings["Sheet1"] = BuildDefaultMappings("Sheet1")
                .Concat(new[] { CreateMappingRow("OtherSheet", "other_sheet_owner", "single", false, currentSingle: "其他负责人") })
                .ToArray();
            var preview = new AiColumnMappingPreview
            {
                Items = new[]
                {
                    new AiColumnMappingPreviewItem
                    {
                        ExcelColumn = 2,
                        SuggestedExcelL1 = "基础信息",
                        SuggestedExcelL2 = "负责人",
                        TargetHeaderId = "owner_name",
                        TargetApiFieldKey = "owner_name",
                        Confidence = 0.91,
                        Status = AiColumnMappingPreviewStatuses.Accepted,
                    },
                },
            };

            var (service, _) = CreateService(connector, metadataStore, new FakeWorksheetSelectionReader(), new FakeAiColumnMappingClient());
            var result = InvokeApplyAiColumnMappingPreview(service, "Sheet1", preview);

            Assert.Equal(1, result.AppliedCount);
            Assert.Equal("基础信息", metadataStore.LastSavedFieldMappings.Single(row => row.Values["HeaderId"] == "owner_name").Values["CurrentL1"]);
            Assert.Equal("负责人", metadataStore.LastSavedFieldMappings.Single(row => row.Values["HeaderId"] == "owner_name").Values["CurrentL2"]);
            Assert.Equal("负责人", metadataStore.LastSavedFieldMappings.Single(row => row.Values["HeaderId"] == "owner_name").Values["DefaultL1"]);
            Assert.Equal(string.Empty, metadataStore.LastSavedFieldMappings.Single(row => row.Values["HeaderId"] == "owner_name").Values["DefaultL2"]);
            Assert.Equal("其他负责人", metadataStore.LastSavedFieldMappings.Single(row => row.SheetName == "OtherSheet").Values["CurrentL1"]);
        }

        private static (object Service, FakeWorksheetGridAdapter Grid) CreateService(
            FakeSystemConnector connector,
            IWorksheetMetadataStore metadataStore,
            FakeWorksheetSelectionReader selectionReader)
        {
            return CreateService(new[] { connector }, metadataStore, selectionReader);
        }

        private static (object Service, FakeWorksheetGridAdapter Grid) CreateService(
            FakeSystemConnector connector,
            IWorksheetMetadataStore metadataStore,
            FakeWorksheetSelectionReader selectionReader,
            IAiColumnMappingClient aiClient)
        {
            return CreateService(new[] { connector }, metadataStore, selectionReader, aiClient);
        }

        private static (object Service, FakeWorksheetGridAdapter Grid) CreateService(
            IReadOnlyList<FakeSystemConnector> connectors,
            IWorksheetMetadataStore metadataStore,
            FakeWorksheetSelectionReader selectionReader,
            IAiColumnMappingClient aiClient = null)
        {
            var assembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var serviceType = assembly.GetType("OfficeAgent.ExcelAddIn.WorksheetSyncExecutionService", throwOnError: true);
            var gridInterface = assembly.GetType("OfficeAgent.ExcelAddIn.Excel.IWorksheetGridAdapter", throwOnError: true);
            var grid = new FakeWorksheetGridAdapter(gridInterface);
            var syncService = new WorksheetSyncService(
                new SystemConnectorRegistry(connectors.Cast<ISystemConnector>().ToArray()),
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

            if (aiClient != null)
            {
                var extendedCtor = serviceType.GetConstructor(
                    BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                    binder: null,
                    types: new[]
                    {
                        typeof(WorksheetSyncService),
                        typeof(IWorksheetMetadataStore),
                        typeof(IWorksheetSelectionReader),
                        gridInterface,
                        typeof(SyncOperationPreviewFactory),
                        typeof(IAiColumnMappingClient),
                    },
                    modifiers: null);

                if (extendedCtor == null)
                {
                    throw new InvalidOperationException("WorksheetSyncExecutionService AI constructor was not found.");
                }

                service = extendedCtor.Invoke(new object[]
                {
                    syncService,
                    metadataStore,
                    selectionReader,
                    grid.GetTransparentProxy(),
                    new SyncOperationPreviewFactory(),
                    aiClient,
                });
            }

            return (service, grid);
        }

        private static (object Service, FakeWorksheetGridAdapter Grid, FakeWorksheetChangeLogStore LogStore, object PendingEditTracker) CreateServiceWithChangeLog(
            FakeSystemConnector connector,
            IWorksheetMetadataStore metadataStore,
            FakeWorksheetSelectionReader selectionReader)
        {
            return CreateServiceWithChangeLog(new[] { connector }, metadataStore, selectionReader);
        }

        private static (object Service, FakeWorksheetGridAdapter Grid, FakeWorksheetChangeLogStore LogStore, object PendingEditTracker) CreateServiceWithChangeLog(
            IReadOnlyList<FakeSystemConnector> connectors,
            IWorksheetMetadataStore metadataStore,
            FakeWorksheetSelectionReader selectionReader)
        {
            var assembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var serviceType = assembly.GetType("OfficeAgent.ExcelAddIn.WorksheetSyncExecutionService", throwOnError: true);
            var gridInterface = assembly.GetType("OfficeAgent.ExcelAddIn.Excel.IWorksheetGridAdapter", throwOnError: true);
            var logStoreInterface = assembly.GetType("OfficeAgent.ExcelAddIn.Excel.IWorksheetChangeLogStore", throwOnError: true);
            var pendingEditTrackerType = assembly.GetType("OfficeAgent.ExcelAddIn.Excel.WorksheetPendingEditTracker", throwOnError: true);
            var grid = new FakeWorksheetGridAdapter(gridInterface);
            var logStore = new FakeWorksheetChangeLogStore(logStoreInterface);
            var pendingEditTracker = Activator.CreateInstance(pendingEditTrackerType);
            var syncService = new WorksheetSyncService(
                new SystemConnectorRegistry(connectors.Cast<ISystemConnector>().ToArray()),
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
                    logStoreInterface,
                    pendingEditTrackerType,
                },
                modifiers: null);

            if (ctor == null)
            {
                throw new InvalidOperationException("WorksheetSyncExecutionService change-log constructor was not found.");
            }

            var service = ctor.Invoke(new object[]
            {
                syncService,
                metadataStore,
                selectionReader,
                grid.GetTransparentProxy(),
                new SyncOperationPreviewFactory(),
                logStore.GetTransparentProxy(),
                pendingEditTracker,
            });

            return (service, grid, logStore, pendingEditTracker);
        }

        private static ScopedWorksheetMetadataAdapter CreateScopedMetadataAdapter()
        {
            var assembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var adapterInterface = assembly.GetType("OfficeAgent.ExcelAddIn.Excel.IWorksheetMetadataAdapter", throwOnError: true);
            return new ScopedWorksheetMetadataAdapter(adapterInterface);
        }

        private static IWorksheetMetadataStore CreateRealMetadataStore(ScopedWorksheetMetadataAdapter adapter)
        {
            var assembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var storeType = assembly.GetType("OfficeAgent.ExcelAddIn.Excel.WorksheetMetadataStore", throwOnError: true);
            var adapterInterface = assembly.GetType("OfficeAgent.ExcelAddIn.Excel.IWorksheetMetadataAdapter", throwOnError: true);
            var ctor = storeType.GetConstructor(
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                binder: null,
                types: new[] { adapterInterface },
                modifiers: null);

            if (ctor == null)
            {
                throw new InvalidOperationException("WorksheetMetadataStore constructor was not found.");
            }

            return (IWorksheetMetadataStore)ctor.Invoke(new[] { adapter.GetTransparentProxy() });
        }

        private static void SeedWorkbookMetadata(
            ScopedWorksheetMetadataAdapter adapter,
            string workbookScopeKey,
            SheetBinding binding,
            IReadOnlyList<SheetFieldMappingRow> mappings)
        {
            var metadataStore = CreateRealMetadataStore(adapter);
            adapter.SwitchWorkbook(workbookScopeKey);
            metadataStore.SaveBinding(binding);
            metadataStore.SaveFieldMappings(
                binding?.SheetName ?? string.Empty,
                BuildDefinition(),
                mappings ?? Array.Empty<SheetFieldMappingRow>());
        }

        private static void SeedRecognizedHeaders(FakeWorksheetGridAdapter grid, string sheetName, SheetBinding binding)
        {
            var row = binding.HeaderStartRow;
            grid.SetCell(sheetName, row, 1, "ID");
            grid.SetCell(sheetName, row, 2, "项目负责人");
            grid.SetCell(sheetName, row, 3, "测试活动111");

            if (binding.HeaderRowCount > 1)
            {
                grid.SetCell(sheetName, row + 1, 3, "开始时间");
                grid.SetCell(sheetName, row + 1, 4, "结束时间");
            }
        }

        private static void SeedGroupedSingleRecognizedHeaders(FakeWorksheetGridAdapter grid, string sheetName, SheetBinding binding)
        {
            var row = binding.HeaderStartRow;
            grid.SetCell(sheetName, row, 1, "ID");
            grid.SetCell(sheetName, row, 2, "联系人信息");
            grid.SetCell(sheetName, row, 3, "测试活动111");

            if (binding.HeaderRowCount > 1)
            {
                grid.SetCell(sheetName, row + 1, 2, "负责人");
                grid.SetCell(sheetName, row + 1, 3, "开始时间");
                grid.SetCell(sheetName, row + 1, 4, "结束时间");
            }
        }

        private static void SeedPendingOriginalValue(object pendingEditTracker, string sheetName, int row, int column, string value)
        {
            var assembly = pendingEditTracker.GetType().Assembly;
            var cellValueType = assembly.GetType("OfficeAgent.ExcelAddIn.Excel.WorksheetCellValue", throwOnError: true);
            var cellAddressType = assembly.GetType("OfficeAgent.ExcelAddIn.Excel.WorksheetCellAddress", throwOnError: true);

            var cellValues = Array.CreateInstance(cellValueType, 1);
            var cellValue = Activator.CreateInstance(cellValueType);
            SetProperty(cellValue, "Row", row);
            SetProperty(cellValue, "Column", column);
            SetProperty(cellValue, "Text", value);
            cellValues.SetValue(cellValue, 0);

            var cellAddresses = Array.CreateInstance(cellAddressType, 1);
            var cellAddress = Activator.CreateInstance(cellAddressType);
            SetProperty(cellAddress, "Row", row);
            SetProperty(cellAddress, "Column", column);
            cellAddresses.SetValue(cellAddress, 0);

            pendingEditTracker.GetType()
                .GetMethod("CaptureBeforeValues")
                .Invoke(pendingEditTracker, new object[] { sheetName, cellValues });
            pendingEditTracker.GetType()
                .GetMethod("MarkChanged")
                .Invoke(pendingEditTracker, new object[] { sheetName, cellAddresses });
        }

        private static bool TryGetPendingOriginalValue(object pendingEditTracker, string sheetName, int row, int column, out string value)
        {
            var args = new object[] { sheetName, row, column, null };
            var result = (bool)pendingEditTracker.GetType()
                .GetMethod("TryGetOriginalValue")
                .Invoke(pendingEditTracker, args);
            value = Convert.ToString(args[3]) ?? string.Empty;
            return result;
        }

        private static void SetProperty(object target, string propertyName, object value)
        {
            target.GetType()
                .GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
                .SetValue(target, value);
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

        private static void InvokeInitialize(object service, string sheetName, ProjectOption project)
        {
            var method = service.GetType().GetMethod(
                "InitializeCurrentSheet",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

            if (method == null)
            {
                throw new InvalidOperationException("InitializeCurrentSheet was not found.");
            }

            method.Invoke(service, new object[] { sheetName, project });
        }

        private static void InvokeTryAutoInitialize(object service, string sheetName, ProjectOption project)
        {
            var method = service.GetType().GetMethod(
                "TryAutoInitializeCurrentSheet",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

            if (method == null)
            {
                throw new InvalidOperationException("TryAutoInitializeCurrentSheet was not found.");
            }

            method.Invoke(service, new object[] { sheetName, project });
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

        private static AiColumnMappingApplyResult InvokeApplyAiColumnMappingPreview(
            object service,
            string sheetName,
            AiColumnMappingPreview preview)
        {
            var method = service.GetType().GetMethod(
                "ApplyAiColumnMappingPreview",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

            if (method == null)
            {
                throw new InvalidOperationException("ApplyAiColumnMappingPreview was not found.");
            }

            try
            {
                return (AiColumnMappingApplyResult)method.Invoke(service, new object[] { sheetName, preview });
            }
            catch (TargetInvocationException ex) when (ex.InnerException != null)
            {
                throw ex.InnerException;
            }
        }

        private static async Task<AiColumnMappingPreview> InvokePrepareAiColumnMappingPreviewAsync(
            object service,
            string sheetName,
            CancellationToken cancellationToken)
        {
            var method = service.GetType().GetMethod(
                "PrepareAiColumnMappingPreviewAsync",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

            if (method == null)
            {
                throw new InvalidOperationException("PrepareAiColumnMappingPreviewAsync was not found.");
            }

            try
            {
                var task = (Task<AiColumnMappingPreview>)method.Invoke(service, new object[] { sheetName, cancellationToken });
                return await task;
            }
            catch (TargetInvocationException ex) when (ex.InnerException != null)
            {
                throw ex.InnerException;
            }
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

        private static bool ReadBoolProperty(object target, string propertyName)
        {
            var property = target.GetType().GetProperty(
                propertyName,
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

            if (property == null)
            {
                throw new InvalidOperationException($"{propertyName} property was not found.");
            }

            return (bool)(property.GetValue(target) ?? false);
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

        private static FieldMappingTableDefinition BuildDefinition()
        {
            return new FieldMappingTableDefinition
            {
                SystemKey = "current-business-system",
                Columns = new[]
                {
                    new FieldMappingColumnDefinition { ColumnName = "HeaderId", Role = FieldMappingSemanticRole.HeaderIdentity },
                    new FieldMappingColumnDefinition { ColumnName = "HeaderType", Role = FieldMappingSemanticRole.HeaderType },
                    new FieldMappingColumnDefinition { ColumnName = "ApiFieldKey", Role = FieldMappingSemanticRole.ApiFieldKey },
                    new FieldMappingColumnDefinition { ColumnName = "IsIdColumn", Role = FieldMappingSemanticRole.IsIdColumn },
                    new FieldMappingColumnDefinition { ColumnName = "ISDP L1", Role = FieldMappingSemanticRole.DefaultSingleHeaderText, RoleKey = "DefaultL1" },
                    new FieldMappingColumnDefinition { ColumnName = "Excel L1", Role = FieldMappingSemanticRole.CurrentSingleHeaderText, RoleKey = "CurrentL1" },
                    new FieldMappingColumnDefinition { ColumnName = "ISDP L1", Role = FieldMappingSemanticRole.DefaultParentHeaderText, RoleKey = "DefaultL1" },
                    new FieldMappingColumnDefinition { ColumnName = "Excel L1", Role = FieldMappingSemanticRole.CurrentParentHeaderText, RoleKey = "CurrentL1" },
                    new FieldMappingColumnDefinition { ColumnName = "ISDP L2", Role = FieldMappingSemanticRole.DefaultChildHeaderText, RoleKey = "DefaultL2" },
                    new FieldMappingColumnDefinition { ColumnName = "Excel L2", Role = FieldMappingSemanticRole.CurrentChildHeaderText, RoleKey = "CurrentL2" },
                    new FieldMappingColumnDefinition { ColumnName = "ActivityId", Role = FieldMappingSemanticRole.ActivityIdentity },
                    new FieldMappingColumnDefinition { ColumnName = "PropertyId", Role = FieldMappingSemanticRole.PropertyIdentity },
                },
            };
        }

        private static SheetFieldMappingRow[] BuildDefaultMappings(string sheetName)
        {
            return new[]
            {
                CreateMappingRow(sheetName, "row_id", "single", true, currentSingle: "ID"),
                CreateMappingRow(sheetName, "owner_name", "single", false, defaultSingle: "负责人", currentSingle: "项目负责人"),
                CreateMappingRow(
                    sheetName,
                    "start_12345678",
                    "activityProperty",
                    false,
                    defaultParent: "测试活动111",
                    currentParent: "测试活动111",
                    defaultChild: "开始时间",
                    currentChild: "开始时间",
                    activityId: "12345678",
                    propertyId: "start"),
                CreateMappingRow(
                    sheetName,
                    "end_12345678",
                    "activityProperty",
                    false,
                    defaultParent: "测试活动111",
                    currentParent: "测试活动111",
                    defaultChild: "结束时间",
                    currentChild: "结束时间",
                    activityId: "12345678",
                    propertyId: "end"),
            };
        }

        private static SheetFieldMappingRow[] BuildSingleRowActivityL1Mappings(string sheetName)
        {
            return new[]
            {
                CreateMappingRow(sheetName, "row_id", "single", true, currentSingle: "ID"),
                CreateMappingRow(sheetName, "owner_name", "single", false, defaultSingle: "负责人", currentSingle: "项目负责人"),
                CreateMappingRow(
                    sheetName,
                    "start_12345678",
                    "activityProperty",
                    false,
                    defaultParent: "测试活动111",
                    currentParent: "计划开始",
                    defaultChild: "开始时间",
                    activityId: "12345678",
                    propertyId: "start"),
                CreateMappingRow(
                    sheetName,
                    "end_12345678",
                    "activityProperty",
                    false,
                    defaultParent: "测试活动111",
                    currentParent: "计划结束",
                    defaultChild: "结束时间",
                    activityId: "12345678",
                    propertyId: "end"),
            };
        }

        private static SheetFieldMappingRow[] BuildLegacyMappingsWithoutIdFlag(string sheetName)
        {
            return new[]
            {
                new SheetFieldMappingRow
                {
                    SheetName = sheetName,
                    Values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["HeaderId"] = "row_id",
                        ["CurrentL1"] = "ID",
                    },
                },
                new SheetFieldMappingRow
                {
                    SheetName = sheetName,
                    Values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["HeaderId"] = "owner_name",
                        ["CurrentL1"] = "项目负责人",
                    },
                },
            };
        }

        private static SheetFieldMappingRow[] BuildGroupedSingleOwnerMappings(string sheetName)
        {
            return new[]
            {
                CreateMappingRow(sheetName, "row_id", "single", true, currentSingle: "ID"),
                CreateMappingRow(
                    sheetName,
                    "owner_name",
                    "single",
                    false,
                    defaultParent: "联系人信息",
                    currentParent: "联系人信息",
                    defaultChild: "负责人",
                    currentChild: "负责人"),
                CreateMappingRow(
                    sheetName,
                    "start_12345678",
                    "activityProperty",
                    false,
                    defaultParent: "测试活动111",
                    currentParent: "测试活动111",
                    defaultChild: "开始时间",
                    currentChild: "开始时间",
                    activityId: "12345678",
                    propertyId: "start"),
                CreateMappingRow(
                    sheetName,
                    "end_12345678",
                    "activityProperty",
                    false,
                    defaultParent: "测试活动111",
                    currentParent: "测试活动111",
                    defaultChild: "结束时间",
                    currentChild: "结束时间",
                    activityId: "12345678",
                    propertyId: "end"),
            };
        }

        private static SheetFieldMappingRow CreateMappingRow(
            string sheetName,
            string apiFieldKey,
            string headerType,
            bool isIdColumn,
            string defaultSingle = "",
            string currentSingle = "",
            string defaultParent = "",
            string currentParent = "",
            string defaultChild = "",
            string currentChild = "",
            string activityId = "",
            string propertyId = "")
        {
            return new SheetFieldMappingRow
            {
                SheetName = sheetName,
                Values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["HeaderId"] = apiFieldKey,
                    ["HeaderType"] = headerType,
                    ["ApiFieldKey"] = apiFieldKey,
                    ["IsIdColumn"] = isIdColumn ? "true" : "false",
                    ["DefaultL1"] = string.IsNullOrWhiteSpace(defaultSingle) ? defaultParent : defaultSingle,
                    ["CurrentL1"] = string.IsNullOrWhiteSpace(currentSingle) ? currentParent : currentSingle,
                    ["DefaultL2"] = defaultChild,
                    ["CurrentL2"] = currentChild,
                    ["ActivityId"] = activityId,
                    ["PropertyId"] = propertyId,
                },
            };
        }

        private static IDictionary<string, object> CreateRow(string rowId, string ownerName, string start, string end)
        {
            return new Dictionary<string, object>(StringComparer.Ordinal)
            {
                ["row_id"] = rowId,
                ["owner_name"] = ownerName,
                ["start_12345678"] = start,
                ["end_12345678"] = end,
            };
        }

        private sealed class FakeSystemConnector : ISystemConnector, IUploadChangeFilter
        {
            public FakeSystemConnector(string systemKey = "current-business-system")
            {
                SystemKey = systemKey;
                BindingSeed = new SheetBinding
                {
                    SheetName = "Sheet1",
                    SystemKey = systemKey,
                    ProjectId = "performance",
                    ProjectName = "绩效项目",
                    HeaderStartRow = 1,
                    HeaderRowCount = 2,
                    DataStartRow = 3,
                };
                FieldMappingDefinition = BuildDefinition();
                FieldMappingSeedRows = BuildDefaultMappings("Sheet1");
            }

            public string SystemKey { get; }

            public SheetBinding BindingSeed { get; set; }

            public FieldMappingTableDefinition FieldMappingDefinition { get; set; }

            public IReadOnlyList<SheetFieldMappingRow> FieldMappingSeedRows { get; set; }

            public IReadOnlyList<IDictionary<string, object>> FindResult { get; set; } = Array.Empty<IDictionary<string, object>>();

            public ProjectOption LastCreateBindingSeedProject { get; private set; }

            public string LastFieldMappingDefinitionProjectId { get; private set; }

            public string LastFindProjectId { get; private set; }

            public IReadOnlyList<string> LastFindRowIds { get; private set; } = Array.Empty<string>();

            public IReadOnlyList<string> LastFindFieldKeys { get; private set; } = Array.Empty<string>();

            public string LastBatchSaveProjectId { get; private set; }

            public IReadOnlyList<CellChange> LastBatchSaveChanges { get; private set; } = Array.Empty<CellChange>();

            public Exception BatchSaveException { get; set; }

            public string SkippedApiFieldKey { get; set; } = string.Empty;

            public string SkipReason { get; set; } = string.Empty;

            public string LastFilterProjectId { get; private set; }

            public IReadOnlyList<ProjectOption> GetProjects()
            {
                return Array.Empty<ProjectOption>();
            }

            public SheetBinding CreateBindingSeed(string sheetName, ProjectOption project)
            {
                LastCreateBindingSeedProject = project;
                return new SheetBinding
                {
                    SheetName = sheetName,
                    SystemKey = project?.SystemKey ?? SystemKey,
                    ProjectId = project?.ProjectId ?? string.Empty,
                    ProjectName = project?.DisplayName ?? string.Empty,
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
                LastFindRowIds = rowIds?.ToArray() ?? Array.Empty<string>();
                LastFindFieldKeys = fieldKeys?.ToArray() ?? Array.Empty<string>();

                IEnumerable<IDictionary<string, object>> rows = FindResult;

                if (LastFindRowIds.Count > 0)
                {
                    rows = rows.Where(row => LastFindRowIds.Contains(Convert.ToString(row["row_id"])));
                }

                if (LastFindFieldKeys.Count > 0)
                {
                    rows = rows.Select(row =>
                    {
                        var projected = new Dictionary<string, object>(StringComparer.Ordinal)
                        {
                            ["row_id"] = row["row_id"],
                        };

                        foreach (var fieldKey in LastFindFieldKeys)
                        {
                            if (row.TryGetValue(fieldKey, out var value))
                            {
                                projected[fieldKey] = value;
                            }
                        }

                        return (IDictionary<string, object>)projected;
                    });
                }

                return rows.ToArray();
            }

            public void BatchSave(string projectId, IReadOnlyList<CellChange> changes)
            {
                if (BatchSaveException != null)
                {
                    throw BatchSaveException;
                }

                LastBatchSaveProjectId = projectId;
                LastBatchSaveChanges = changes?.ToArray() ?? Array.Empty<CellChange>();
            }

            public UploadChangeFilterResult FilterUploadChanges(string projectId, IReadOnlyList<CellChange> changes)
            {
                LastFilterProjectId = projectId;
                var changeList = changes ?? Array.Empty<CellChange>();

                return new UploadChangeFilterResult
                {
                    IncludedChanges = changeList
                        .Where(change => !string.Equals(change.ApiFieldKey, SkippedApiFieldKey, StringComparison.Ordinal))
                        .ToArray(),
                    SkippedChanges = changeList
                        .Where(change => string.Equals(change.ApiFieldKey, SkippedApiFieldKey, StringComparison.Ordinal))
                        .Select(change => new SkippedCellChange
                        {
                            Change = change,
                            Reason = SkipReason,
                        })
                        .ToArray(),
                };
            }
        }

        private sealed class FakeAiColumnMappingClient : IAiColumnMappingClient
        {
            public AiColumnMappingResponse Response { get; set; } = new AiColumnMappingResponse();

            public AiColumnMappingRequest LastRequest { get; private set; }

            public CancellationToken LastCancellationToken { get; private set; }

            public AiColumnMappingResponse Map(AiColumnMappingRequest request)
            {
                LastRequest = request;
                return Response;
            }

            public System.Threading.Tasks.Task<AiColumnMappingResponse> MapAsync(AiColumnMappingRequest request)
            {
                return MapAsync(request, CancellationToken.None);
            }

            public System.Threading.Tasks.Task<AiColumnMappingResponse> MapAsync(AiColumnMappingRequest request, CancellationToken cancellationToken)
            {
                LastRequest = request;
                LastCancellationToken = cancellationToken;
                return System.Threading.Tasks.Task.FromResult(Response);
            }
        }

        private sealed class FakeWorksheetChangeLogStore : RealProxy
        {
            private readonly Type interfaceType;

            public FakeWorksheetChangeLogStore(Type interfaceType)
                : base(interfaceType)
            {
                this.interfaceType = interfaceType;
            }

            public List<ChangeLogRecord> Entries { get; } = new List<ChangeLogRecord>();

            public override IMessage Invoke(IMessage msg)
            {
                var call = (IMethodCallMessage)msg;

                switch (call.MethodName)
                {
                    case "GetType":
                        return new ReturnMessage(interfaceType, null, 0, call.LogicalCallContext, call);
                    case "GetHashCode":
                        return new ReturnMessage(GetHashCode(), null, 0, call.LogicalCallContext, call);
                    case "ToString":
                        return new ReturnMessage(nameof(FakeWorksheetChangeLogStore), null, 0, call.LogicalCallContext, call);
                    case "Append":
                        Append((IEnumerable)call.InArgs[0]);
                        return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
                    default:
                        throw new NotSupportedException(call.MethodName);
                }
            }

            public new object GetTransparentProxy()
            {
                return base.GetTransparentProxy();
            }

            private void Append(IEnumerable entries)
            {
                foreach (var entry in entries ?? Array.Empty<object>())
                {
                    Entries.Add(new ChangeLogRecord
                    {
                        Key = ReadString(entry, "Key"),
                        HeaderText = ReadString(entry, "HeaderText"),
                        ChangeMode = ReadString(entry, "ChangeMode"),
                        NewValue = ReadString(entry, "NewValue"),
                        OldValue = ReadString(entry, "OldValue"),
                    });
                }
            }

            private static string ReadString(object target, string propertyName)
            {
                return Convert.ToString(target.GetType().GetProperty(propertyName).GetValue(target)) ?? string.Empty;
            }
        }

        private sealed class ChangeLogRecord
        {
            public string Key { get; set; } = string.Empty;

            public string HeaderText { get; set; } = string.Empty;

            public string ChangeMode { get; set; } = string.Empty;

            public string NewValue { get; set; } = string.Empty;

            public string OldValue { get; set; } = string.Empty;
        }

        private sealed class FakeWorksheetMetadataStore : IWorksheetMetadataStore
        {
            public Dictionary<string, SheetBinding> Bindings { get; } = new Dictionary<string, SheetBinding>(StringComparer.OrdinalIgnoreCase);

            public Dictionary<string, SheetFieldMappingRow[]> FieldMappings { get; } = new Dictionary<string, SheetFieldMappingRow[]>(StringComparer.OrdinalIgnoreCase);

            public SheetBinding LastSavedBinding { get; private set; }

            public FieldMappingTableDefinition LastSavedFieldMappingDefinition { get; private set; }

            public SheetFieldMappingRow[] LastSavedFieldMappings { get; private set; } = Array.Empty<SheetFieldMappingRow>();

            public void SaveBinding(SheetBinding binding)
            {
                LastSavedBinding = binding;
                Bindings[binding.SheetName] = binding;
            }

            public void RefreshMetadataPresentation(string sheetName, bool hideTemplateBindingRows = false)
            {
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

        private sealed class FakeWorksheetSelectionReader : IWorksheetSelectionReader
        {
            public IReadOnlyList<SelectedVisibleCell> VisibleCells { get; set; } = Array.Empty<SelectedVisibleCell>();

            public WorksheetSelectionSnapshot SelectionSnapshot { get; set; } = new WorksheetSelectionSnapshot();

            public bool ThrowOnReadVisibleSelection { get; set; }

            public IReadOnlyList<SelectedVisibleCell> ReadVisibleSelection()
            {
                if (ThrowOnReadVisibleSelection)
                {
                    throw new InvalidOperationException("Visible cell enumeration should not be used for large selection downloads.");
                }

                return VisibleCells;
            }

            public WorksheetSelectionSnapshot ReadSelectionSnapshot()
            {
                return SelectionSnapshot;
            }
        }

        private sealed class FakeWorksheetGridAdapter : RealProxy
        {
            private readonly Dictionary<string, FakeCell> cells = new Dictionary<string, FakeCell>(StringComparer.OrdinalIgnoreCase);
            private int bulkOperationDepth;

            public FakeWorksheetGridAdapter(Type interfaceType)
                : base(interfaceType)
            {
            }

            public List<MergeRecord> Merges { get; } = new List<MergeRecord>();

            public List<ClearRangeRecord> ClearedRanges { get; } = new List<ClearRangeRecord>();

            public List<WriteRangeRecord> WriteRangeCalls { get; } = new List<WriteRangeRecord>();

            public List<ReadRangeRecord> ReadRangeCalls { get; } = new List<ReadRangeRecord>();

            public List<GetCellTextRecord> GetCellTextCalls { get; } = new List<GetCellTextRecord>();

            public List<LastUsedRowRecord> LastUsedRowCalls { get; } = new List<LastUsedRowRecord>();

            public int BeginBulkOperationCount { get; private set; }

            public int EndBulkOperationCount { get; private set; }

            public override IMessage Invoke(IMessage msg)
            {
                var call = (IMethodCallMessage)msg;

                switch (call.MethodName)
                {
                    case "GetCellText":
                        {
                            var sheetName = (string)call.InArgs[0];
                            var row = (int)call.InArgs[1];
                            var column = (int)call.InArgs[2];
                            GetCellTextCalls.Add(new GetCellTextRecord
                            {
                                SheetName = sheetName,
                                Row = row,
                                Column = column,
                            });
                            return new ReturnMessage(GetCell(sheetName, row, column), null, 0, call.LogicalCallContext, call);
                        }
                    case "SetCellText":
                        SetCell(
                            (string)call.InArgs[0],
                            (int)call.InArgs[1],
                            (int)call.InArgs[2],
                            (string)call.InArgs[3]);
                        return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
                    case "ClearRange":
                        ClearRange(
                            (string)call.InArgs[0],
                            (int)call.InArgs[1],
                            (int)call.InArgs[2],
                            (int)call.InArgs[3],
                            (int)call.InArgs[4]);
                        return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
                    case "ClearWorksheet":
                        ClearWorksheet((string)call.InArgs[0]);
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
                    case "GetLastUsedColumn":
                        return new ReturnMessage(GetLastUsedColumn((string)call.InArgs[0]), null, 0, call.LogicalCallContext, call);
                    case "WriteRangeValues":
                        WriteRangeValues(
                            (string)call.InArgs[0],
                            (int)call.InArgs[1],
                            (int)call.InArgs[2],
                            (object[,])call.InArgs[3]);
                        return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
                    case "ReadRangeValues":
                        {
                            var sheetName = (string)call.InArgs[0];
                            var startRow = (int)call.InArgs[1];
                            var endRow = (int)call.InArgs[2];
                            var startColumn = (int)call.InArgs[3];
                            var endColumn = (int)call.InArgs[4];
                            ReadRangeCalls.Add(new ReadRangeRecord
                            {
                                MethodName = "ReadRangeValues",
                                SheetName = sheetName,
                                StartRow = startRow,
                                EndRow = endRow,
                                StartColumn = startColumn,
                                EndColumn = endColumn,
                                WasInsideBulkOperation = IsBulkOperationActive,
                            });
                            return new ReturnMessage(
                                ReadRangeValues(sheetName, startRow, endRow, startColumn, endColumn),
                                null,
                                0,
                                call.LogicalCallContext,
                                call);
                        }
                    case "ReadRangeNumberFormats":
                        {
                            var sheetName = (string)call.InArgs[0];
                            var startRow = (int)call.InArgs[1];
                            var endRow = (int)call.InArgs[2];
                            var startColumn = (int)call.InArgs[3];
                            var endColumn = (int)call.InArgs[4];
                            ReadRangeCalls.Add(new ReadRangeRecord
                            {
                                MethodName = "ReadRangeNumberFormats",
                                SheetName = sheetName,
                                StartRow = startRow,
                                EndRow = endRow,
                                StartColumn = startColumn,
                                EndColumn = endColumn,
                                WasInsideBulkOperation = IsBulkOperationActive,
                            });
                            return new ReturnMessage(
                                ReadRangeNumberFormats(sheetName, startRow, endRow, startColumn, endColumn),
                                null,
                                0,
                                call.LogicalCallContext,
                                call);
                        }
                    case "BeginBulkOperation":
                        BeginBulkOperationCount++;
                        bulkOperationDepth++;
                        return new ReturnMessage(
                            new DelegateDisposeScope(() =>
                            {
                                if (bulkOperationDepth > 0)
                                {
                                    bulkOperationDepth--;
                                }

                                EndBulkOperationCount++;
                            }),
                            null,
                            0,
                            call.LogicalCallContext,
                            call);
                    default:
                        throw new NotSupportedException(call.MethodName);
                }
            }

            public void SetCell(string sheetName, int row, int column, string value)
            {
                cells[BuildKey(sheetName, row, column)] = new FakeCell
                {
                    Text = value ?? string.Empty,
                    RawValue = value ?? string.Empty,
                    NumberFormat = string.Empty,
                };
            }

            public void SetRawCell(string sheetName, int row, int column, object rawValue, string numberFormat = "", string text = null)
            {
                cells[BuildKey(sheetName, row, column)] = new FakeCell
                {
                    Text = text ?? Convert.ToString(rawValue) ?? string.Empty,
                    RawValue = rawValue,
                    NumberFormat = numberFormat ?? string.Empty,
                };
            }

            public string GetCell(string sheetName, int row, int column)
            {
                return cells.TryGetValue(BuildKey(sheetName, row, column), out var cell)
                    ? cell.Text
                    : string.Empty;
            }

            public void ClearAllCells()
            {
                cells.Clear();
            }

            public int CountGetCellTextCalls(string sheetName, int row, int column)
            {
                return GetCellTextCalls.Count(call =>
                    string.Equals(call.SheetName, sheetName, StringComparison.OrdinalIgnoreCase) &&
                    call.Row == row &&
                    call.Column == column);
            }

            public new object GetTransparentProxy()
            {
                return base.GetTransparentProxy();
            }

            private void ClearRange(string sheetName, int startRow, int endRow, int startColumn, int endColumn)
            {
                ClearedRanges.Add(new ClearRangeRecord
                {
                    SheetName = sheetName,
                    StartRow = startRow,
                    EndRow = endRow,
                    StartColumn = startColumn,
                    EndColumn = endColumn,
                });

                var keysToRemove = cells.Keys
                    .Where(key => IsWithinRange(key, sheetName, startRow, endRow, startColumn, endColumn))
                    .ToArray();

                foreach (var key in keysToRemove)
                {
                    cells.Remove(key);
                }
            }

            private void ClearWorksheet(string sheetName)
            {
                var keysToRemove = cells.Keys
                    .Where(key => key.StartsWith(sheetName + "|", StringComparison.OrdinalIgnoreCase))
                    .ToArray();

                foreach (var key in keysToRemove)
                {
                    cells.Remove(key);
                }
            }

            private int GetLastUsedRow(string sheetName)
            {
                LastUsedRowCalls.Add(new LastUsedRowRecord
                {
                    SheetName = sheetName,
                    WasInsideBulkOperation = IsBulkOperationActive,
                });

                var prefix = sheetName + "|";
                var rows = cells.Keys
                    .Where(key => key.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    .Select(key => int.Parse(key.Split('|')[1]))
                    .ToArray();

                return rows.Length == 0 ? 0 : rows.Max();
            }

            private int GetLastUsedColumn(string sheetName)
            {
                var prefix = sheetName + "|";
                var columns = cells.Keys
                    .Where(key => key.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    .Select(key => int.Parse(key.Split('|')[2]))
                    .ToArray();

                return columns.Length == 0 ? 0 : columns.Max();
            }

            private void WriteRangeValues(string sheetName, int startRow, int startColumn, object[,] values)
            {
                WriteRangeCalls.Add(new WriteRangeRecord
                {
                    SheetName = sheetName,
                    StartRow = startRow,
                    StartColumn = startColumn,
                    Values = values,
                    WasInsideBulkOperation = IsBulkOperationActive,
                });

                if (values == null)
                {
                    return;
                }

                for (var rowOffset = 0; rowOffset < values.GetLength(0); rowOffset++)
                {
                    for (var columnOffset = 0; columnOffset < values.GetLength(1); columnOffset++)
                    {
                        SetRawCell(
                            sheetName,
                            startRow + rowOffset,
                            startColumn + columnOffset,
                            values[rowOffset, columnOffset],
                            text: Convert.ToString(values[rowOffset, columnOffset]) ?? string.Empty);
                    }
                }
            }

            private object[,] ReadRangeValues(
                string sheetName,
                int startRow,
                int endRow,
                int startColumn,
                int endColumn)
            {
                var rowCount = Math.Max(0, endRow - startRow + 1);
                var columnCount = Math.Max(0, endColumn - startColumn + 1);
                var values = new object[rowCount, columnCount];
                for (var rowOffset = 0; rowOffset < rowCount; rowOffset++)
                {
                    for (var columnOffset = 0; columnOffset < columnCount; columnOffset++)
                    {
                        values[rowOffset, columnOffset] = cells.TryGetValue(
                            BuildKey(sheetName, startRow + rowOffset, startColumn + columnOffset),
                            out var cell)
                            ? cell.RawValue
                            : string.Empty;
                    }
                }

                return values;
            }

            private string[,] ReadRangeNumberFormats(
                string sheetName,
                int startRow,
                int endRow,
                int startColumn,
                int endColumn)
            {
                var rowCount = Math.Max(0, endRow - startRow + 1);
                var columnCount = Math.Max(0, endColumn - startColumn + 1);
                var formats = new string[rowCount, columnCount];
                for (var rowOffset = 0; rowOffset < rowCount; rowOffset++)
                {
                    for (var columnOffset = 0; columnOffset < columnCount; columnOffset++)
                    {
                        formats[rowOffset, columnOffset] = cells.TryGetValue(
                            BuildKey(sheetName, startRow + rowOffset, startColumn + columnOffset),
                            out var cell)
                            ? cell.NumberFormat
                            : string.Empty;
                    }
                }

                return formats;
            }

            private static bool IsWithinRange(
                string key,
                string sheetName,
                int startRow,
                int endRow,
                int startColumn,
                int endColumn)
            {
                var parts = key.Split('|');
                if (parts.Length != 3)
                {
                    return false;
                }

                if (!string.Equals(parts[0], sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                var row = int.Parse(parts[1]);
                var column = int.Parse(parts[2]);
                return row >= startRow &&
                       row <= endRow &&
                       column >= startColumn &&
                       column <= endColumn;
            }

            private static string BuildKey(string sheetName, int row, int column)
            {
                return string.Join("|", sheetName ?? string.Empty, row, column);
            }

            private bool IsBulkOperationActive => bulkOperationDepth > 0;

            private sealed class FakeCell
            {
                public string Text { get; set; } = string.Empty;

                public object RawValue { get; set; } = string.Empty;

                public string NumberFormat { get; set; } = string.Empty;
            }
        }

        private sealed class ScopedWorksheetMetadataAdapter : RealProxy
        {
            private readonly Dictionary<string, Dictionary<string, List<string[]>>> tablesByWorkbook =
                new Dictionary<string, Dictionary<string, List<string[]>>>(StringComparer.OrdinalIgnoreCase);
            private readonly Dictionary<string, Dictionary<string, string[]>> headersByWorkbook =
                new Dictionary<string, Dictionary<string, string[]>>(StringComparer.OrdinalIgnoreCase);

            public ScopedWorksheetMetadataAdapter(Type adapterInterface)
                : base(adapterInterface)
            {
            }

            public string WorkbookScopeKey { get; private set; } = "Workbook1";

            public override IMessage Invoke(IMessage msg)
            {
                var call = (IMethodCallMessage)msg;

                switch (call.MethodName)
                {
                    case "GetWorkbookScopeKey":
                        return new ReturnMessage(WorkbookScopeKey, null, 0, call.LogicalCallContext, call);
                    case "EnsureWorksheet":
                        return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
                    case "ApplyMetadataPresentation":
                        return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
                    case "WriteTable":
                        {
                            var tables = GetCurrentWorkbookTables();
                            var headers = GetCurrentWorkbookHeaders();
                            var tableName = (string)call.InArgs[0];
                            var tableHeaders = (string[])call.InArgs[1];
                            var rows = (string[][])call.InArgs[2];
                            headers[tableName] = tableHeaders?.ToArray() ?? Array.Empty<string>();
                            tables[tableName] = (rows ?? Array.Empty<string[]>())
                                .Select(row => row?.ToArray() ?? Array.Empty<string>())
                                .ToList();
                            return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
                        }
                    case "ReadHeaders":
                        {
                            var headers = GetCurrentWorkbookHeaders();
                            var tableName = (string)call.InArgs[0];
                            var result = headers.TryGetValue(tableName, out var storedHeaders)
                                ? storedHeaders.ToArray()
                                : Array.Empty<string>();
                            return new ReturnMessage(result, null, 0, call.LogicalCallContext, call);
                        }
                    case "ReadTable":
                        {
                            var tables = GetCurrentWorkbookTables();
                            var tableName = (string)call.InArgs[0];
                            var rows = tables.TryGetValue(tableName, out var storedRows)
                                ? storedRows.Select(row => row.ToArray()).ToArray()
                                : Array.Empty<string[]>();
                            return new ReturnMessage(rows, null, 0, call.LogicalCallContext, call);
                        }
                    default:
                        throw new NotSupportedException(call.MethodName);
                }
            }

            public void SwitchWorkbook(string workbookScopeKey)
            {
                WorkbookScopeKey = workbookScopeKey ?? string.Empty;
                GetCurrentWorkbookTables();
                GetCurrentWorkbookHeaders();
            }

            public void SeedTable(string tableName, string[][] rows)
            {
                var tables = GetCurrentWorkbookTables();
                tables[tableName] = (rows ?? Array.Empty<string[]>())
                    .Select(row => row?.ToArray() ?? Array.Empty<string>())
                    .ToList();
            }

            public new object GetTransparentProxy()
            {
                return base.GetTransparentProxy();
            }

            private Dictionary<string, List<string[]>> GetCurrentWorkbookTables()
            {
                if (!tablesByWorkbook.TryGetValue(WorkbookScopeKey, out var tables))
                {
                    tables = new Dictionary<string, List<string[]>>(StringComparer.OrdinalIgnoreCase);
                    tablesByWorkbook[WorkbookScopeKey] = tables;
                }

                return tables;
            }

            private Dictionary<string, string[]> GetCurrentWorkbookHeaders()
            {
                if (!headersByWorkbook.TryGetValue(WorkbookScopeKey, out var headers))
                {
                    headers = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);
                    headersByWorkbook[WorkbookScopeKey] = headers;
                }

                return headers;
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

        public sealed class ClearRangeRecord
        {
            public string SheetName { get; set; }
            public int StartRow { get; set; }
            public int EndRow { get; set; }
            public int StartColumn { get; set; }
            public int EndColumn { get; set; }
        }

        public sealed class WriteRangeRecord
        {
            public string SheetName { get; set; }
            public int StartRow { get; set; }
            public int StartColumn { get; set; }
            public object[,] Values { get; set; }
            public bool WasInsideBulkOperation { get; set; }
        }

        public sealed class ReadRangeRecord
        {
            public string MethodName { get; set; }
            public string SheetName { get; set; }
            public int StartRow { get; set; }
            public int EndRow { get; set; }
            public int StartColumn { get; set; }
            public int EndColumn { get; set; }
            public bool WasInsideBulkOperation { get; set; }
        }

        public sealed class LastUsedRowRecord
        {
            public string SheetName { get; set; }
            public bool WasInsideBulkOperation { get; set; }
        }

        public sealed class GetCellTextRecord
        {
            public string SheetName { get; set; }
            public int Row { get; set; }
            public int Column { get; set; }
        }

        private sealed class DelegateDisposeScope : IDisposable
        {
            private readonly Action onDispose;
            private bool disposed;

            public DelegateDisposeScope(Action onDispose)
            {
                this.onDispose = onDispose;
            }

            public void Dispose()
            {
                if (disposed)
                {
                    return;
                }

                disposed = true;
                onDispose?.Invoke();
            }
        }
    }
}
