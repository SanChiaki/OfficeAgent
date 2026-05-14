using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Sync;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class WorksheetHeaderMatcherTests
    {
        [Fact]
        public void MatchUsesParentAndChildDisplayNamesForTwoRowHeaders()
        {
            var grid = new FakeGrid();
            grid.SetCell("Sheet1", 3, 1, "ID");
            grid.SetCell("Sheet1", 3, 2, "测试活动111");
            grid.SetCell("Sheet1", 4, 2, "开始时间");
            grid.SetCell("Sheet1", 4, 3, "结束时间");

            var matcher = CreateMatcher();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
            };
            var definition = BuildDefinition();
            var mappings = BuildActivityMappings("Sheet1");

            var columns = InvokeMatch(matcher, "Sheet1", binding, definition, mappings, grid);

            Assert.Contains(columns, column => column.ColumnIndex == 2 && column.ApiFieldKey == "start_12345678");
            Assert.Contains(columns, column => column.ColumnIndex == 3 && column.ApiFieldKey == "end_12345678");
        }

        [Fact]
        public void MatchUsesSingleDisplayNameForSingleRowHeaders()
        {
            var grid = new FakeGrid();
            grid.SetCell("Sheet1", 5, 1, "ID");
            grid.SetCell("Sheet1", 5, 2, "项目负责人");

            var matcher = CreateMatcher();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                HeaderStartRow = 5,
                HeaderRowCount = 1,
            };
            var definition = BuildDefinition();
            var mappings = new[]
            {
                CreateMappingRow(
                    "Sheet1",
                    apiFieldKey: "row_id",
                    headerType: "single",
                    isIdColumn: true,
                    currentSingle: "ID"),
                CreateMappingRow(
                    "Sheet1",
                    apiFieldKey: "owner_name",
                    headerType: "single",
                    isIdColumn: false,
                    currentSingle: "项目负责人"),
            };

            var columns = InvokeMatch(matcher, "Sheet1", binding, definition, mappings, grid);

            Assert.Contains(columns, column => column.ColumnIndex == 1 && column.ApiFieldKey == "row_id" && column.IsIdColumn);
            Assert.Contains(columns, column => column.ColumnIndex == 2 && column.ApiFieldKey == "owner_name");
            Assert.Equal(2, columns.Length);
        }

        [Fact]
        public void MatchUsesActivityExcelL1ForSingleRowHeaders()
        {
            var grid = new FakeGrid();
            grid.SetCell("Sheet1", 5, 1, "ID");
            grid.SetCell("Sheet1", 5, 2, "计划开始");
            grid.SetCell("Sheet1", 5, 3, "计划结束");

            var matcher = CreateMatcher();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                HeaderStartRow = 5,
                HeaderRowCount = 1,
            };
            var definition = BuildDefinition();
            var mappings = new[]
            {
                CreateMappingRow(
                    "Sheet1",
                    apiFieldKey: "row_id",
                    headerType: "single",
                    isIdColumn: true,
                    currentSingle: "ID"),
                CreateMappingRow(
                    "Sheet1",
                    apiFieldKey: "start_12345678",
                    headerType: "activityProperty",
                    isIdColumn: false,
                    currentParent: "计划开始"),
                CreateMappingRow(
                    "Sheet1",
                    apiFieldKey: "end_12345678",
                    headerType: "activityProperty",
                    isIdColumn: false,
                    currentParent: "计划结束"),
            };

            var columns = InvokeMatch(matcher, "Sheet1", binding, definition, mappings, grid);

            Assert.Contains(columns, column =>
                column.ColumnIndex == 2 &&
                column.ApiFieldKey == "start_12345678" &&
                column.HeaderType == "activityProperty" &&
                column.DisplayText == "计划开始");
            Assert.Contains(columns, column =>
                column.ColumnIndex == 3 &&
                column.ApiFieldKey == "end_12345678" &&
                column.DisplayText == "计划结束");
            Assert.Equal(3, columns.Length);
        }

        [Fact]
        public void MatchUsesGroupedSingleDisplayNamesForTwoRowHeaders()
        {
            var grid = new FakeGrid();
            grid.SetCell("Sheet1", 3, 1, "ID");
            grid.SetCell("Sheet1", 3, 2, "基础信息");
            grid.SetCell("Sheet1", 4, 2, "负责人");

            var matcher = CreateMatcher();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
            };
            var definition = BuildDefinition();
            var mappings = new[]
            {
                CreateMappingRow(
                    "Sheet1",
                    apiFieldKey: "row_id",
                    headerType: "single",
                    isIdColumn: true,
                    currentSingle: "ID"),
                CreateMappingRow(
                    "Sheet1",
                    apiFieldKey: "owner_name",
                    headerType: "single",
                    isIdColumn: false,
                    currentParent: "基础信息",
                    currentChild: "负责人"),
            };

            var columns = InvokeMatch(matcher, "Sheet1", binding, definition, mappings, grid);

            Assert.Contains(columns, column => column.ColumnIndex == 2 && column.ApiFieldKey == "owner_name");
        }

        [Fact]
        public void MatchPrefersMergedNormalSingleBeforeTwoLevelMatching()
        {
            var grid = new FakeGrid();
            grid.SetCell("Sheet1", 3, 1, "负责人");
            grid.SetCell("Sheet1", 4, 1, "负责人");

            var matcher = CreateMatcher();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
            };
            var definition = BuildDefinition();
            var mappings = new[]
            {
                CreateMappingRow(
                    "Sheet1",
                    apiFieldKey: "owner_name",
                    headerType: "single",
                    isIdColumn: false,
                    currentSingle: "负责人"),
                CreateMappingRow(
                    "Sheet1",
                    apiFieldKey: "owner_grouped",
                    headerType: "single",
                    isIdColumn: false,
                    currentParent: "负责人",
                    currentChild: "负责人"),
            };

            var columns = InvokeMatch(matcher, "Sheet1", binding, definition, mappings, grid);

            var matched = Assert.Single(columns);
            Assert.Equal("owner_name", matched.ApiFieldKey);
        }

        [Fact]
        public void MatchThrowsWhenGroupedSingleConflictsWithActivityTwoLevelKey()
        {
            var grid = new FakeGrid();
            grid.SetCell("Sheet1", 3, 1, "测试活动111");
            grid.SetCell("Sheet1", 4, 1, "开始时间");

            var matcher = CreateMatcher();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                HeaderStartRow = 3,
                HeaderRowCount = 2,
            };
            var definition = BuildDefinition();
            var mappings = new[]
            {
                CreateMappingRow(
                    "Sheet1",
                    apiFieldKey: "start_grouped",
                    headerType: "single",
                    isIdColumn: false,
                    currentParent: "测试活动111",
                    currentChild: "开始时间"),
                CreateMappingRow(
                    "Sheet1",
                    apiFieldKey: "start_activity",
                    headerType: "activityProperty",
                    isIdColumn: false,
                    currentParent: "测试活动111",
                    currentChild: "开始时间"),
            };

            var exception = Assert.Throws<InvalidOperationException>(
                () => InvokeMatch(matcher, "Sheet1", binding, definition, mappings, grid));

            Assert.Equal(
                "SheetFieldMappings 中存在重复的双层表头键，请先修正 xISDP_Setting。",
                exception.Message);
        }

        [Fact]
        public void MatchThrowsWhenSingleWithExcelL2UsedInOneRowHeader()
        {
            var grid = new FakeGrid();
            grid.SetCell("Sheet1", 5, 1, "负责人");

            var matcher = CreateMatcher();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                HeaderStartRow = 5,
                HeaderRowCount = 1,
            };
            var definition = BuildDefinition();
            var mappings = new[]
            {
                CreateMappingRow(
                    "Sheet1",
                    apiFieldKey: "owner_grouped",
                    headerType: "single",
                    isIdColumn: false,
                    currentParent: "基础信息",
                    currentChild: "负责人"),
            };

            var exception = Assert.Throws<InvalidOperationException>(
                () => InvokeMatch(matcher, "Sheet1", binding, definition, mappings, grid));

            Assert.Equal(
                "当前 HeaderRowCount=1，无法识别带 Excel L2 的 single 表头，请先修正 xISDP_Setting。",
                exception.Message);
        }

        [Fact]
        public void MatchIgnoresGroupedSingleMetadataFromOtherSheetWhenHeaderRowCountIsOne()
        {
            var grid = new FakeGrid();
            grid.SetCell("Sheet1", 5, 1, "ID");

            var matcher = CreateMatcher();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                HeaderStartRow = 5,
                HeaderRowCount = 1,
            };
            var definition = BuildDefinition();
            var mappings = new[]
            {
                CreateMappingRow(
                    "Sheet1",
                    apiFieldKey: "row_id",
                    headerType: "single",
                    isIdColumn: true,
                    currentSingle: "ID"),
                CreateMappingRow(
                    "Sheet2",
                    apiFieldKey: "owner_grouped",
                    headerType: "single",
                    isIdColumn: false,
                    currentParent: "基础信息",
                    currentChild: "负责人"),
            };

            var columns = InvokeMatch(matcher, "Sheet1", binding, definition, mappings, grid);

            var matched = Assert.Single(columns);
            Assert.Equal("row_id", matched.ApiFieldKey);
        }

        private static object CreateMatcher()
        {
            var assembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var matcherType = assembly.GetType("OfficeAgent.ExcelAddIn.Excel.WorksheetHeaderMatcher", throwOnError: true);
            var constructor = matcherType.GetConstructor(
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                binder: null,
                types: new[] { typeof(FieldMappingValueAccessor) },
                modifiers: null);
            return constructor.Invoke(new object[] { new FieldMappingValueAccessor() });
        }

        private static WorksheetRuntimeColumn[] InvokeMatch(
            object matcher,
            string sheetName,
            SheetBinding binding,
            FieldMappingTableDefinition definition,
            IReadOnlyList<SheetFieldMappingRow> mappings,
            FakeGrid grid)
        {
            var method = matcher.GetType().GetMethod(
                "Match",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            try
            {
                return (WorksheetRuntimeColumn[])method.Invoke(
                    matcher,
                    new object[] { sheetName, binding, definition, mappings, grid.GetTransparentProxy() });
            }
            catch (TargetInvocationException ex) when (ex.InnerException != null)
            {
                throw ex.InnerException;
            }
        }

        private static FieldMappingTableDefinition BuildDefinition()
        {
            return new FieldMappingTableDefinition
            {
                SystemKey = "current-business-system",
                Columns = new[]
                {
                    new FieldMappingColumnDefinition { ColumnName = "ApiFieldKey", Role = FieldMappingSemanticRole.ApiFieldKey },
                    new FieldMappingColumnDefinition { ColumnName = "HeaderType", Role = FieldMappingSemanticRole.HeaderType },
                    new FieldMappingColumnDefinition { ColumnName = "IsIdColumn", Role = FieldMappingSemanticRole.IsIdColumn },
                    new FieldMappingColumnDefinition { ColumnName = "Excel L1", Role = FieldMappingSemanticRole.CurrentSingleHeaderText, RoleKey = "CurrentL1" },
                    new FieldMappingColumnDefinition { ColumnName = "Excel L1", Role = FieldMappingSemanticRole.CurrentParentHeaderText, RoleKey = "CurrentL1" },
                    new FieldMappingColumnDefinition { ColumnName = "Excel L2", Role = FieldMappingSemanticRole.CurrentChildHeaderText, RoleKey = "CurrentL2" },
                },
            };
        }

        private static SheetFieldMappingRow[] BuildActivityMappings(string sheetName)
        {
            return new[]
            {
                CreateMappingRow(
                    sheetName,
                    apiFieldKey: "row_id",
                    headerType: "single",
                    isIdColumn: true,
                    currentSingle: "ID"),
                CreateMappingRow(
                    sheetName,
                    apiFieldKey: "start_12345678",
                    headerType: "activityProperty",
                    isIdColumn: false,
                    currentParent: "测试活动111",
                    currentChild: "开始时间"),
                CreateMappingRow(
                    sheetName,
                    apiFieldKey: "end_12345678",
                    headerType: "activityProperty",
                    isIdColumn: false,
                    currentParent: "测试活动111",
                    currentChild: "结束时间"),
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
                    ["CurrentL1"] = string.IsNullOrWhiteSpace(currentSingle) ? currentParent : currentSingle,
                    ["CurrentL2"] = currentChild,
                },
            };
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

        private sealed class FakeGrid : RealProxy
        {
            private readonly Dictionary<(string Sheet, int Row, int Column), string> cells =
                new Dictionary<(string Sheet, int Row, int Column), string>();

            public FakeGrid()
                : base(GetAdapterType())
            {
            }

            public void SetCell(string sheetName, int row, int column, string value)
            {
                cells[(sheetName, row, column)] = value;
            }

            public override IMessage Invoke(IMessage msg)
            {
                var call = (IMethodCallMessage)msg;
                switch (call.MethodName)
                {
                    case "GetCellText":
                        return HandleGetCellText(call);
                    case "GetLastUsedColumn":
                        return HandleGetLastUsedColumn(call);
                    case "GetLastUsedRow":
                        return new ReturnMessage(0, null, 0, call.LogicalCallContext, call);
                    case "SetCellText":
                    case "ClearWorksheet":
                    case "ClearRange":
                    case "MergeCells":
                        return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
                    default:
                        throw new NotSupportedException(call.MethodName);
                }
            }

            private IMessage HandleGetCellText(IMethodCallMessage call)
            {
                var sheetName = (string)call.InArgs[0];
                var row = (int)call.InArgs[1];
                var column = (int)call.InArgs[2];
                cells.TryGetValue((sheetName, row, column), out var value);
                return new ReturnMessage(value ?? string.Empty, null, 0, call.LogicalCallContext, call);
            }

            private IMessage HandleGetLastUsedColumn(IMethodCallMessage call)
            {
                var sheetName = (string)call.InArgs[0];
                var lastColumn = cells.Keys
                    .Where(key => string.Equals(key.Sheet, sheetName, StringComparison.OrdinalIgnoreCase))
                    .Select(key => key.Column)
                    .DefaultIfEmpty(0)
                    .Max();
                return new ReturnMessage(lastColumn, null, 0, call.LogicalCallContext, call);
            }

            private static Type GetAdapterType()
            {
                return Assembly.LoadFrom(ResolveAddInAssemblyPath())
                    .GetType("OfficeAgent.ExcelAddIn.Excel.IWorksheetGridAdapter", throwOnError: true);
            }
        }
    }
}
