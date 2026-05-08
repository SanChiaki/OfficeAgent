using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using OfficeAgent.Core.Models;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class WorksheetHeaderScannerTests
    {
        [Fact]
        public void ScanReadsAllUsedColumnsInOneRowHeaderArea()
        {
            var scanner = CreateScanner();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                HeaderStartRow = 3,
                HeaderRowCount = 1,
            };
            var grid = new FakeGrid();
            grid.SetCell("Sheet1", 3, 1, "ID");
            grid.SetCell("Sheet1", 3, 3, "项目负责人");
            grid.SetCell("Sheet1", 6, 5, "data extends used range");

            var headers = InvokeScan(scanner, "Sheet1", binding, grid);

            Assert.Collection(
                headers,
                header =>
                {
                    Assert.Equal(1, header.ExcelColumn);
                    Assert.Equal("ID", header.ActualL1);
                    Assert.Equal(string.Empty, header.ActualL2);
                    Assert.Equal("ID", header.DisplayText);
                },
                header =>
                {
                    Assert.Equal(3, header.ExcelColumn);
                    Assert.Equal("项目负责人", header.ActualL1);
                    Assert.Equal(string.Empty, header.ActualL2);
                    Assert.Equal("项目负责人", header.DisplayText);
                });
        }

        [Fact]
        public void ScanCarriesParentTextAcrossTwoRowHeaderArea()
        {
            var scanner = CreateScanner();
            var binding = new SheetBinding
            {
                SheetName = "Sheet1",
                HeaderStartRow = 4,
                HeaderRowCount = 2,
            };
            var grid = new FakeGrid();
            grid.SetCell("Sheet1", 4, 1, "ID");
            grid.SetCell("Sheet1", 4, 2, "基础信息");
            grid.SetCell("Sheet1", 5, 2, "负责人");
            grid.SetCell("Sheet1", 4, 3, "测试活动111");
            grid.SetCell("Sheet1", 5, 3, "开始时间");
            grid.SetCell("Sheet1", 5, 4, "结束时间");

            var headers = InvokeScan(scanner, "Sheet1", binding, grid);

            Assert.Collection(
                headers,
                header =>
                {
                    Assert.Equal(1, header.ExcelColumn);
                    Assert.Equal("ID", header.ActualL1);
                    Assert.Equal(string.Empty, header.ActualL2);
                },
                header =>
                {
                    Assert.Equal(2, header.ExcelColumn);
                    Assert.Equal("基础信息", header.ActualL1);
                    Assert.Equal("负责人", header.ActualL2);
                    Assert.Equal("基础信息/负责人", header.DisplayText);
                },
                header =>
                {
                    Assert.Equal(3, header.ExcelColumn);
                    Assert.Equal("测试活动111", header.ActualL1);
                    Assert.Equal("开始时间", header.ActualL2);
                },
                header =>
                {
                    Assert.Equal(4, header.ExcelColumn);
                    Assert.Equal("测试活动111", header.ActualL1);
                    Assert.Equal("结束时间", header.ActualL2);
                    Assert.Equal("测试活动111/结束时间", header.DisplayText);
                });
        }

        private static object CreateScanner()
        {
            var assembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var scannerType = assembly.GetType("OfficeAgent.ExcelAddIn.Excel.WorksheetHeaderScanner", throwOnError: true);
            return Activator.CreateInstance(scannerType);
        }

        private static AiColumnMappingActualHeader[] InvokeScan(
            object scanner,
            string sheetName,
            SheetBinding binding,
            FakeGrid grid)
        {
            var method = scanner.GetType().GetMethod(
                "Scan",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            try
            {
                return (AiColumnMappingActualHeader[])method.Invoke(
                    scanner,
                    new[] { sheetName, binding, grid.GetTransparentProxy() });
            }
            catch (TargetInvocationException ex) when (ex.InnerException != null)
            {
                throw ex.InnerException;
            }
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
                    case "WriteRangeValues":
                    case "ClearWorksheet":
                    case "ClearRange":
                    case "MergeCells":
                        return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
                    default:
                        throw new NotSupportedException(call.MethodName);
                }
            }

            public new object GetTransparentProxy()
            {
                return base.GetTransparentProxy();
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
