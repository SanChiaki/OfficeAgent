using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class WorksheetChangeLogStoreTests
    {
        [Fact]
        public void AppendCreatesLogSheetAndKeepsLatestTwoThousandRows()
        {
            var assembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var gridInterface = assembly.GetType("OfficeAgent.ExcelAddIn.Excel.IWorksheetGridAdapter", throwOnError: true);
            var storeType = assembly.GetType("OfficeAgent.ExcelAddIn.Excel.WorksheetChangeLogStore", throwOnError: true);
            var entryType = assembly.GetType("OfficeAgent.ExcelAddIn.Excel.WorksheetChangeLogEntry", throwOnError: true);
            var grid = new FakeWorksheetGridAdapter(gridInterface);
            var store = Activator.CreateInstance(
                storeType,
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                binder: null,
                args: new object[]
                {
                    grid.GetTransparentProxy(),
                    new Func<DateTime>(() => new DateTime(2026, 4, 29, 9, 30, 0)),
                },
                culture: null);

            var entries = Array.CreateInstance(entryType, 2001);
            for (var index = 0; index < entries.Length; index++)
            {
                var entry = Activator.CreateInstance(entryType);
                SetProperty(entry, "Key", $"row-{index + 1:0000}");
                SetProperty(entry, "HeaderText", $"表头{index + 1}");
                SetProperty(entry, "ChangeMode", "下载");
                SetProperty(entry, "NewValue", $"新值{index + 1}");
                SetProperty(entry, "OldValue", $"旧值{index + 1}");
                entries.SetValue(entry, index);
            }

            storeType.GetMethod("Append").Invoke(store, new object[] { entries });

            Assert.Contains("xISDP_Log", grid.WorksheetNames);
            Assert.Equal("key", grid.GetCell("xISDP_Log", 1, 1));
            Assert.Equal("表头", grid.GetCell("xISDP_Log", 1, 2));
            Assert.Equal("修改模式", grid.GetCell("xISDP_Log", 1, 3));
            Assert.Equal("修改值", grid.GetCell("xISDP_Log", 1, 4));
            Assert.Equal("原始值", grid.GetCell("xISDP_Log", 1, 5));
            Assert.Equal("修改时间", grid.GetCell("xISDP_Log", 1, 6));
            Assert.Equal("row-0002", grid.GetCell("xISDP_Log", 2, 1));
            Assert.Equal("表头2", grid.GetCell("xISDP_Log", 2, 2));
            Assert.Equal("下载", grid.GetCell("xISDP_Log", 2, 3));
            Assert.Equal("新值2", grid.GetCell("xISDP_Log", 2, 4));
            Assert.Equal("旧值2", grid.GetCell("xISDP_Log", 2, 5));
            Assert.Equal("2026-04-29 09:30:00", grid.GetCell("xISDP_Log", 2, 6));
            Assert.Equal("row-2001", grid.GetCell("xISDP_Log", 2001, 1));
            Assert.Equal(2001, grid.GetLastUsedRow("xISDP_Log"));
        }

        private static void SetProperty(object target, string propertyName, object value)
        {
            target.GetType()
                .GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
                .SetValue(target, value);
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

        private sealed class FakeWorksheetGridAdapter : RealProxy
        {
            private readonly Type interfaceType;
            private readonly Dictionary<string, string> cells = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            public FakeWorksheetGridAdapter(Type interfaceType)
                : base(interfaceType)
            {
                this.interfaceType = interfaceType;
            }

            public HashSet<string> WorksheetNames { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

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
                        return new ReturnMessage(nameof(FakeWorksheetGridAdapter), null, 0, call.LogicalCallContext, call);
                    case "EnsureWorksheetExists":
                        WorksheetNames.Add((string)call.InArgs[0]);
                        return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
                    case "GetLastUsedRow":
                        return new ReturnMessage(GetLastUsedRow((string)call.InArgs[0]), null, 0, call.LogicalCallContext, call);
                    case "ReadRangeValues":
                        return new ReturnMessage(
                            ReadRangeValues(
                                (string)call.InArgs[0],
                                (int)call.InArgs[1],
                                (int)call.InArgs[2],
                                (int)call.InArgs[3],
                                (int)call.InArgs[4]),
                            null,
                            0,
                            call.LogicalCallContext,
                            call);
                    case "ClearRange":
                        ClearRange(
                            (string)call.InArgs[0],
                            (int)call.InArgs[1],
                            (int)call.InArgs[2],
                            (int)call.InArgs[3],
                            (int)call.InArgs[4]);
                        return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
                    case "WriteRangeValues":
                        WriteRangeValues(
                            (string)call.InArgs[0],
                            (int)call.InArgs[1],
                            (int)call.InArgs[2],
                            (object[,])call.InArgs[3]);
                        return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
                    case "BeginBulkOperation":
                        return new ReturnMessage(new DelegateDisposeScope(), null, 0, call.LogicalCallContext, call);
                    default:
                        throw new NotSupportedException(call.MethodName);
                }
            }

            public new object GetTransparentProxy()
            {
                return base.GetTransparentProxy();
            }

            public string GetCell(string sheetName, int row, int column)
            {
                return cells.TryGetValue(BuildKey(sheetName, row, column), out var value)
                    ? value
                    : string.Empty;
            }

            public int GetLastUsedRow(string sheetName)
            {
                var prefix = sheetName + "|";
                var rows = cells.Keys
                    .Where(key => key.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    .Select(key => int.Parse(key.Split('|')[1]))
                    .ToArray();

                return rows.Length == 0 ? 0 : rows.Max();
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
                        values[rowOffset, columnOffset] = GetCell(
                            sheetName,
                            startRow + rowOffset,
                            startColumn + columnOffset);
                    }
                }

                return values;
            }

            private void ClearRange(string sheetName, int startRow, int endRow, int startColumn, int endColumn)
            {
                var keys = cells.Keys
                    .Where(key => IsWithinRange(key, sheetName, startRow, endRow, startColumn, endColumn))
                    .ToArray();

                foreach (var key in keys)
                {
                    cells.Remove(key);
                }
            }

            private void WriteRangeValues(string sheetName, int startRow, int startColumn, object[,] values)
            {
                if (values == null)
                {
                    return;
                }

                WorksheetNames.Add(sheetName);
                for (var rowOffset = 0; rowOffset < values.GetLength(0); rowOffset++)
                {
                    for (var columnOffset = 0; columnOffset < values.GetLength(1); columnOffset++)
                    {
                        cells[BuildKey(sheetName, startRow + rowOffset, startColumn + columnOffset)] =
                            Convert.ToString(values[rowOffset, columnOffset]) ?? string.Empty;
                    }
                }
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
                return string.Equals(parts[0], sheetName, StringComparison.OrdinalIgnoreCase) &&
                       int.Parse(parts[1]) >= startRow &&
                       int.Parse(parts[1]) <= endRow &&
                       int.Parse(parts[2]) >= startColumn &&
                       int.Parse(parts[2]) <= endColumn;
            }

            private static string BuildKey(string sheetName, int row, int column)
            {
                return $"{sheetName}|{row}|{column}";
            }
        }

        private sealed class DelegateDisposeScope : IDisposable
        {
            public void Dispose()
            {
            }
        }
    }
}
