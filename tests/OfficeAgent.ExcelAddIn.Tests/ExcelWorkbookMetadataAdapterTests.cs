using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class ExcelWorkbookMetadataAdapterTests
    {
        [Fact]
        public void EnsureWorksheetRestoresTheOriginalActiveWorksheetWhenMetadataSheetIsCreated()
        {
            var addInAssembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var excelAssembly = LoadExcelInteropAssembly();
            var worksheetType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Worksheet", throwOnError: true);
            var sheetsType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Sheets", throwOnError: true);
            var workbookType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Workbook", throwOnError: true);
            var applicationType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Application", throwOnError: true);
            var rangeType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Range", throwOnError: true);

            var application = new LayoutAwareFakeExcelApplication(applicationType, workbookType, sheetsType, worksheetType, rangeType);
            var adapterType = addInAssembly.GetType("OfficeAgent.ExcelAddIn.Excel.ExcelWorkbookMetadataAdapter", throwOnError: true);
            var adapter = Activator.CreateInstance(adapterType, application.GetTransparentProxy());
            var ensureWorksheet = adapterType.GetMethod("EnsureWorksheet", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

            ensureWorksheet.Invoke(adapter, new object[] { "_Settings", true });

            Assert.Equal("BusinessSheet", application.ActiveSheet.Name);
            Assert.NotNull(application.MetadataSheet);
            Assert.Equal("_Settings", application.MetadataSheet.Name);
            Assert.Equal(1, application.BusinessSheet.ActivateCount);
        }

        [Fact]
        public void WriteTableRewritesMetadataSheetUsingTitledSections()
        {
            var addInAssembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var excelAssembly = LoadExcelInteropAssembly();
            var worksheetType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Worksheet", throwOnError: true);
            var sheetsType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Sheets", throwOnError: true);
            var workbookType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Workbook", throwOnError: true);
            var applicationType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Application", throwOnError: true);
            var rangeType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Range", throwOnError: true);

            var application = new LayoutAwareFakeExcelApplication(applicationType, workbookType, sheetsType, worksheetType, rangeType);
            var adapterType = addInAssembly.GetType("OfficeAgent.ExcelAddIn.Excel.ExcelWorkbookMetadataAdapter", throwOnError: true);
            var adapter = Activator.CreateInstance(adapterType, application.GetTransparentProxy());

            adapterType.GetMethod("WriteTable").Invoke(adapter, new object[]
            {
                "SheetBindings",
                new[] { "SheetName", "SystemKey" },
                new[] { new[] { "Sheet1", "current-business-system" } },
            });

            adapterType.GetMethod("WriteTable").Invoke(adapter, new object[]
            {
                "SheetFieldMappings",
                new[] { "SheetName", "HeaderId", "ApiFieldKey" },
                new[] { new[] { "Sheet1", "row_id", "row_id" } },
            });

            Assert.Equal("SheetBindings", application.MetadataSheet.GetCell(1, 1));
            Assert.Equal("SheetName", application.MetadataSheet.GetCell(2, 1));
            Assert.Equal("Sheet1", application.MetadataSheet.GetCell(3, 1));
            Assert.Equal(string.Empty, application.MetadataSheet.GetCell(4, 1));
            Assert.Equal(string.Empty, application.MetadataSheet.GetCell(5, 1));
            Assert.Equal("SheetFieldMappings", application.MetadataSheet.GetCell(6, 1));
            Assert.Equal("HeaderId", application.MetadataSheet.GetCell(7, 2));
            Assert.Equal("row_id", application.MetadataSheet.GetCell(8, 2));
        }

        [Fact]
        public void ReadTableReadsRowsBackFromTitledSections()
        {
            var addInAssembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var excelAssembly = LoadExcelInteropAssembly();
            var worksheetType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Worksheet", throwOnError: true);
            var sheetsType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Sheets", throwOnError: true);
            var workbookType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Workbook", throwOnError: true);
            var applicationType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Application", throwOnError: true);
            var rangeType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Range", throwOnError: true);

            var application = new LayoutAwareFakeExcelApplication(applicationType, workbookType, sheetsType, worksheetType, rangeType);
            application.CreateWorksheet("_Settings");
            application.MetadataSheet.SetCell(1, 1, "SheetBindings");
            application.MetadataSheet.SetCell(2, 1, "SheetName");
            application.MetadataSheet.SetCell(2, 2, "SystemKey");
            application.MetadataSheet.SetCell(3, 1, "Sheet1");
            application.MetadataSheet.SetCell(3, 2, "current-business-system");

            var adapterType = addInAssembly.GetType("OfficeAgent.ExcelAddIn.Excel.ExcelWorkbookMetadataAdapter", throwOnError: true);
            var adapter = Activator.CreateInstance(adapterType, application.GetTransparentProxy());

            var rows = (string[][])adapterType.GetMethod("ReadTable").Invoke(adapter, new object[] { "SheetBindings" });

            Assert.Single(rows);
            Assert.Equal(new[] { "Sheet1", "current-business-system" }, rows[0]);
        }

        private static Assembly LoadExcelInteropAssembly()
        {
            try
            {
                return Assembly.Load("Microsoft.Office.Interop.Excel");
            }
            catch (FileNotFoundException)
            {
                var fallbackPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                    "Microsoft Visual Studio",
                    "Shared",
                    "Visual Studio Tools for Office",
                    "PIA",
                    "Office15",
                    "Microsoft.Office.Interop.Excel.dll");
                return Assembly.LoadFrom(fallbackPath);
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

        private sealed class LayoutAwareFakeExcelApplication : RealProxy
        {
            private readonly LayoutAwareFakeExcelWorkbook workbook;
            private readonly Type applicationType;

            public LayoutAwareFakeExcelApplication(Type applicationType, Type workbookType, Type sheetsType, Type worksheetType, Type rangeType)
                : base(applicationType)
            {
                this.applicationType = applicationType;
                BusinessSheet = new LayoutAwareFakeExcelWorksheet(worksheetType, rangeType, this, "BusinessSheet");
                workbook = new LayoutAwareFakeExcelWorkbook(workbookType, sheetsType, worksheetType, rangeType, this, BusinessSheet);
                ActiveSheet = BusinessSheet;
            }

            public LayoutAwareFakeExcelWorksheet BusinessSheet { get; }

            public LayoutAwareFakeExcelWorksheet MetadataSheet { get; set; }

            public LayoutAwareFakeExcelWorksheet ActiveSheet { get; set; }

            public LayoutAwareFakeExcelWorksheet CreateWorksheet(string name)
            {
                var worksheet = workbook.Worksheets.AddWorksheet(name);
                if (string.Equals(name, "_Settings", StringComparison.OrdinalIgnoreCase))
                {
                    MetadataSheet = worksheet;
                }

                return worksheet;
            }

            public override IMessage Invoke(IMessage msg)
            {
                var call = (IMethodCallMessage)msg;

                return call.MethodName switch
                {
                    "GetType" => new ReturnMessage(applicationType, null, 0, call.LogicalCallContext, call),
                    "get_ActiveWorkbook" => new ReturnMessage(workbook.GetTransparentProxy(), null, 0, call.LogicalCallContext, call),
                    "get_ActiveSheet" => new ReturnMessage(ActiveSheet?.GetTransparentProxy(), null, 0, call.LogicalCallContext, call),
                    _ => throw new NotSupportedException(call.MethodName),
                };
            }

            public new object GetTransparentProxy()
            {
                return base.GetTransparentProxy();
            }
        }

        private sealed class LayoutAwareFakeExcelWorkbook : RealProxy
        {
            private readonly Type workbookType;

            public LayoutAwareFakeExcelWorkbook(
                Type workbookType,
                Type sheetsType,
                Type worksheetType,
                Type rangeType,
                LayoutAwareFakeExcelApplication application,
                LayoutAwareFakeExcelWorksheet initialSheet)
                : base(workbookType)
            {
                this.workbookType = workbookType;
                Worksheets = new LayoutAwareFakeExcelSheets(sheetsType, worksheetType, rangeType, application, initialSheet);
            }

            public LayoutAwareFakeExcelSheets Worksheets { get; }

            public override IMessage Invoke(IMessage msg)
            {
                var call = (IMethodCallMessage)msg;

                return call.MethodName switch
                {
                    "GetType" => new ReturnMessage(workbookType, null, 0, call.LogicalCallContext, call),
                    "get_Worksheets" => new ReturnMessage(Worksheets.GetTransparentProxy(), null, 0, call.LogicalCallContext, call),
                    _ => throw new NotSupportedException(call.MethodName),
                };
            }

            public new object GetTransparentProxy()
            {
                return base.GetTransparentProxy();
            }
        }

        private sealed class LayoutAwareFakeExcelSheets : RealProxy
        {
            private readonly List<LayoutAwareFakeExcelWorksheet> sheets = new List<LayoutAwareFakeExcelWorksheet>();
            private readonly Type sheetsType;
            private readonly Type worksheetType;
            private readonly Type rangeType;
            private readonly LayoutAwareFakeExcelApplication application;

            public LayoutAwareFakeExcelSheets(
                Type sheetsType,
                Type worksheetType,
                Type rangeType,
                LayoutAwareFakeExcelApplication application,
                LayoutAwareFakeExcelWorksheet initialSheet)
                : base(sheetsType)
            {
                this.sheetsType = sheetsType;
                this.worksheetType = worksheetType;
                this.rangeType = rangeType;
                this.application = application;
                sheets.Add(initialSheet);
            }

            public override IMessage Invoke(IMessage msg)
            {
                var call = (IMethodCallMessage)msg;

                return call.MethodName switch
                {
                    "GetType" => new ReturnMessage(sheetsType, null, 0, call.LogicalCallContext, call),
                    "get_Count" => new ReturnMessage(sheets.Count, null, 0, call.LogicalCallContext, call),
                    "get_Item" => new ReturnMessage(ResolveSheet(call.InArgs[0]).GetTransparentProxy(), null, 0, call.LogicalCallContext, call),
                    "get__Default" => new ReturnMessage(ResolveSheet(call.InArgs[0]).GetTransparentProxy(), null, 0, call.LogicalCallContext, call),
                    "Item" => new ReturnMessage(ResolveSheet(call.InArgs[0]).GetTransparentProxy(), null, 0, call.LogicalCallContext, call),
                    "Add" => HandleAdd(call),
                    "GetEnumerator" => new ReturnMessage(GetEnumerator(), null, 0, call.LogicalCallContext, call),
                    _ => throw new NotSupportedException(call.MethodName),
                };
            }

            public LayoutAwareFakeExcelWorksheet AddWorksheet(string name)
            {
                var worksheet = new LayoutAwareFakeExcelWorksheet(worksheetType, rangeType, application, name);
                sheets.Add(worksheet);
                return worksheet;
            }

            private IMessage HandleAdd(IMethodCallMessage call)
            {
                var added = AddWorksheet("Sheet" + (sheets.Count + 1));
                application.MetadataSheet = added;
                application.ActiveSheet = added;
                return new ReturnMessage(added.GetTransparentProxy(), null, 0, call.LogicalCallContext, call);
            }

            private IEnumerator GetEnumerator()
            {
                foreach (var sheet in sheets)
                {
                    yield return sheet.GetTransparentProxy();
                }
            }

            private LayoutAwareFakeExcelWorksheet ResolveSheet(object index)
            {
                var ordinal = Convert.ToInt32(index);
                return sheets[ordinal - 1];
            }

            public new object GetTransparentProxy()
            {
                return base.GetTransparentProxy();
            }
        }

        private sealed class LayoutAwareFakeExcelWorksheet : RealProxy
        {
            private readonly Dictionary<(int Row, int Column), string> cells =
                new Dictionary<(int Row, int Column), string>();
            private readonly LayoutAwareFakeExcelApplication application;
            private readonly Type worksheetType;
            private readonly Type rangeType;

            public LayoutAwareFakeExcelWorksheet(Type worksheetType, Type rangeType, LayoutAwareFakeExcelApplication application, string name)
                : base(worksheetType)
            {
                this.worksheetType = worksheetType;
                this.rangeType = rangeType;
                this.application = application;
                Name = name;
            }

            public string Name { get; private set; }

            public int ActivateCount { get; private set; }

            public void SetCell(int row, int column, string value)
            {
                if (string.IsNullOrEmpty(value))
                {
                    cells.Remove((row, column));
                    return;
                }

                cells[(row, column)] = value;
            }

            public string GetCell(int row, int column)
            {
                return cells.TryGetValue((row, column), out var value) ? value : string.Empty;
            }

            public void ClearAllCells()
            {
                cells.Clear();
            }

            public (int FirstRow, int RowCount, int ColumnCount)? GetUsedRange()
            {
                if (cells.Count == 0)
                {
                    return null;
                }

                var firstRow = cells.Keys.Min(key => key.Row);
                var lastRow = cells.Keys.Max(key => key.Row);
                var lastColumn = cells.Keys.Max(key => key.Column);
                return (firstRow, lastRow - firstRow + 1, lastColumn);
            }

            public override IMessage Invoke(IMessage msg)
            {
                var call = (IMethodCallMessage)msg;

                return call.MethodName switch
                {
                    "GetType" => new ReturnMessage(worksheetType, null, 0, call.LogicalCallContext, call),
                    "get_Name" => new ReturnMessage(Name, null, 0, call.LogicalCallContext, call),
                    "set_Name" => SetName(call),
                    "set_Visible" => new ReturnMessage(null, null, 0, call.LogicalCallContext, call),
                    "Activate" => Activate(call),
                    "get_Cells" => new ReturnMessage(
                        new LayoutAwareFakeExcelRange(rangeType, this, LayoutAwareFakeExcelRangeKind.CellCollection).GetTransparentProxy(),
                        null,
                        0,
                        call.LogicalCallContext,
                        call),
                    "get_UsedRange" => GetUsedRangeMessage(call),
                    _ => throw new NotSupportedException(call.MethodName),
                };
            }

            private IMessage SetName(IMethodCallMessage call)
            {
                Name = Convert.ToString(call.InArgs[0]) ?? string.Empty;
                return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
            }

            private IMessage Activate(IMethodCallMessage call)
            {
                ActivateCount++;
                application.ActiveSheet = this;
                return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
            }

            private IMessage GetUsedRangeMessage(IMethodCallMessage call)
            {
                var usedRange = GetUsedRange();
                if (!usedRange.HasValue)
                {
                    return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
                }

                var range = new LayoutAwareFakeExcelRange(
                    rangeType,
                    this,
                    LayoutAwareFakeExcelRangeKind.UsedRange,
                    row: usedRange.Value.FirstRow,
                    count: usedRange.Value.RowCount,
                    otherCount: usedRange.Value.ColumnCount);
                return new ReturnMessage(range.GetTransparentProxy(), null, 0, call.LogicalCallContext, call);
            }

            public new object GetTransparentProxy()
            {
                return base.GetTransparentProxy();
            }
        }

        private enum LayoutAwareFakeExcelRangeKind
        {
            CellCollection,
            Cell,
            UsedRange,
            Rows,
            Columns,
        }

        private sealed class LayoutAwareFakeExcelRange : RealProxy
        {
            private readonly Type rangeType;
            private readonly LayoutAwareFakeExcelWorksheet worksheet;
            private readonly LayoutAwareFakeExcelRangeKind kind;
            private readonly int row;
            private readonly int column;
            private readonly int count;
            private readonly int otherCount;

            public LayoutAwareFakeExcelRange(
                Type rangeType,
                LayoutAwareFakeExcelWorksheet worksheet,
                LayoutAwareFakeExcelRangeKind kind,
                int row = 0,
                int column = 0,
                int count = 0,
                int otherCount = 0)
                : base(rangeType)
            {
                this.rangeType = rangeType;
                this.worksheet = worksheet;
                this.kind = kind;
                this.row = row;
                this.column = column;
                this.count = count;
                this.otherCount = otherCount;
            }

            public override IMessage Invoke(IMessage msg)
            {
                var call = (IMethodCallMessage)msg;

                return call.MethodName switch
                {
                    "GetType" => new ReturnMessage(rangeType, null, 0, call.LogicalCallContext, call),
                    "get_Item" => HandleGetItem(call),
                    "get__Default" => HandleGetItem(call),
                    "Item" => HandleGetItem(call),
                    "get_Value2" => new ReturnMessage(worksheet.GetCell(row, column), null, 0, call.LogicalCallContext, call),
                    "set_Value2" => HandleSetValue(call),
                    "ClearContents" => HandleClearContents(call),
                    "get_Row" => new ReturnMessage(row, null, 0, call.LogicalCallContext, call),
                    "get_Rows" => new ReturnMessage(
                        new LayoutAwareFakeExcelRange(rangeType, worksheet, LayoutAwareFakeExcelRangeKind.Rows, count: count).GetTransparentProxy(),
                        null,
                        0,
                        call.LogicalCallContext,
                        call),
                    "get_Columns" => new ReturnMessage(
                        new LayoutAwareFakeExcelRange(rangeType, worksheet, LayoutAwareFakeExcelRangeKind.Columns, count: otherCount).GetTransparentProxy(),
                        null,
                        0,
                        call.LogicalCallContext,
                        call),
                    "get_Count" => new ReturnMessage(count, null, 0, call.LogicalCallContext, call),
                    _ => throw new NotSupportedException(call.MethodName),
                };
            }

            private IMessage HandleGetItem(IMethodCallMessage call)
            {
                if (kind != LayoutAwareFakeExcelRangeKind.CellCollection)
                {
                    throw new NotSupportedException(call.MethodName + ":" + kind);
                }

                var cellRow = Convert.ToInt32(call.InArgs[0]);
                var cellColumn = Convert.ToInt32(call.InArgs[1]);
                var cell = new LayoutAwareFakeExcelRange(
                    rangeType,
                    worksheet,
                    LayoutAwareFakeExcelRangeKind.Cell,
                    row: cellRow,
                    column: cellColumn);
                return new ReturnMessage(cell.GetTransparentProxy(), null, 0, call.LogicalCallContext, call);
            }

            private IMessage HandleSetValue(IMethodCallMessage call)
            {
                worksheet.SetCell(row, column, Convert.ToString(call.InArgs[0]) ?? string.Empty);
                return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
            }

            private IMessage HandleClearContents(IMethodCallMessage call)
            {
                if (kind == LayoutAwareFakeExcelRangeKind.CellCollection)
                {
                    worksheet.ClearAllCells();
                }
                else if (kind == LayoutAwareFakeExcelRangeKind.Cell)
                {
                    worksheet.SetCell(row, column, string.Empty);
                }
                else
                {
                    throw new NotSupportedException(call.MethodName + ":" + kind);
                }

                return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
            }

            public new object GetTransparentProxy()
            {
                return base.GetTransparentProxy();
            }
        }
    }
}
