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
    public sealed class WorksheetMetadataStoreTests
    {
        [Fact]
        public void SaveBindingCreatesVisibleMetadataWorksheetAndRoundTripsBinding()
        {
            var (store, adapter) = CreateStore();
            var binding = new SheetBinding
            {
                SheetName = "Sync-performance",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
            };

            InvokeSaveBinding(store, binding);

            Assert.Equal("_OfficeAgentMetadata", adapter.WorksheetName);
            Assert.True(adapter.Visible);

            var loaded = InvokeLoadBinding(store, "Sync-performance");

            Assert.Equal("performance", loaded.ProjectId);
            Assert.Equal("绩效项目", loaded.ProjectName);
        }

        [Fact]
        public void SaveBindingPreservesOtherSheetBindings()
        {
            var (store, adapter) = CreateStore();
            adapter.SeedTable("SheetBindings", new[]
            {
                new[] { "Existing", "system-legacy", "legacy-project", "Legacy" },
            });

            var newBinding = new SheetBinding
            {
                SheetName = "NewSheet",
                SystemKey = "system-new",
                ProjectId = "new-project",
                ProjectName = "New Project",
            };

            InvokeSaveBinding(store, newBinding);

            var legacy = InvokeLoadBinding(store, "Existing");
            Assert.Equal("legacy-project", legacy.ProjectId);

            var added = InvokeLoadBinding(store, "NewSheet");
            Assert.Equal("new-project", added.ProjectId);
        }

        [Fact]
        public void SaveSnapshotPreservesOtherSheetSnapshots()
        {
            var (store, adapter) = CreateStore();
            adapter.SeedTable("SheetSnapshots", new[]
            {
                new[] { "SheetA", "row-1", "field-a", "value-a" },
                new[] { "SheetB", "row-1", "field-b", "value-b" },
            });

            var replacementCells = new[]
            {
                new WorksheetSnapshotCell
                {
                    RowId = "row-2",
                    ApiFieldKey = "field-a",
                    Value = "value-a-new",
                },
            };

            InvokeSaveSnapshot(store, "SheetA", replacementCells);

            var sheetA = InvokeLoadSnapshot(store, "SheetA");
            Assert.Single(sheetA);
            Assert.Equal("row-2", sheetA[0].RowId);
            Assert.Equal("value-a-new", sheetA[0].Value);

            var sheetB = InvokeLoadSnapshot(store, "SheetB");
            Assert.Single(sheetB);
            Assert.Equal("field-b", sheetB[0].ApiFieldKey);
            Assert.Equal("value-b", sheetB[0].Value);
        }

        private static (object Store, FakeWorksheetMetadataAdapter Adapter) CreateStore()
        {
            var assembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var storeType = GetAddInType(assembly, "OfficeAgent.ExcelAddIn.Excel.WorksheetMetadataStore");
            var adapterInterface = GetAddInType(assembly, "OfficeAgent.ExcelAddIn.Excel.IWorksheetMetadataAdapter");
            var adapter = new FakeWorksheetMetadataAdapter(adapterInterface);
            var proxy = adapter.GetTransparentProxy();

            var ctor = storeType.GetConstructor(
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                binder: null,
                types: new[] { adapterInterface },
                modifiers: null);

            var store = ctor.Invoke(new[] { proxy });
            return (store, adapter);
        }

        private static void InvokeSaveBinding(object store, SheetBinding binding)
        {
            var method = store.GetType().GetMethod(
                "SaveBinding",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            method.Invoke(store, new object[] { binding });
        }

        private static SheetBinding InvokeLoadBinding(object store, string sheetName)
        {
            var method = store.GetType().GetMethod(
                "LoadBinding",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            return (SheetBinding)method.Invoke(store, new object[] { sheetName });
        }

        private static void InvokeSaveSnapshot(object store, string sheetName, WorksheetSnapshotCell[] cells)
        {
            var method = store.GetType().GetMethod(
                "SaveSnapshot",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            method.Invoke(store, new object[] { sheetName, cells });
        }

        private static WorksheetSnapshotCell[] InvokeLoadSnapshot(object store, string sheetName)
        {
            var method = store.GetType().GetMethod(
                "LoadSnapshot",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            return (WorksheetSnapshotCell[])method.Invoke(store, new object[] { sheetName });
        }

        private static Type GetAddInType(Assembly assembly, string typeName)
        {
            return assembly.GetType(typeName, throwOnError: true);
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

        private sealed class FakeWorksheetMetadataAdapter : RealProxy
        {
            private readonly Dictionary<string, List<string[]>> tables =
                new Dictionary<string, List<string[]>>(StringComparer.OrdinalIgnoreCase);

            public string WorksheetName { get; private set; }
            public bool Visible { get; private set; }

            public FakeWorksheetMetadataAdapter(Type adapterInterface)
                : base(adapterInterface)
            {
            }

            public override IMessage Invoke(IMessage msg)
            {
                var call = (IMethodCallMessage)msg;

                return call.MethodName switch
                {
                    "EnsureWorksheet" => HandleEnsureWorksheet(call),
                    "WriteTable" => HandleWriteTable(call),
                    "ReadTable" => HandleReadTable(call),
                    _ => throw new NotSupportedException(call.MethodName),
                };
            }

            private IMessage HandleEnsureWorksheet(IMethodCallMessage call)
            {
                WorksheetName = (string)call.InArgs[0];
                Visible = (bool)call.InArgs[1];
                return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
            }

            private IMessage HandleWriteTable(IMethodCallMessage call)
            {
                var tableName = (string)call.InArgs[0];
                var rows = (string[][])call.InArgs[2];
                if (rows == null)
                {
                    tables.Remove(tableName);
                }
                else
                {
                    tables[tableName] = rows.Select(row => row?.ToArray() ?? Array.Empty<string>()).ToList();
                }

                return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
            }

            private IMessage HandleReadTable(IMethodCallMessage call)
            {
                var tableName = (string)call.InArgs[0];
                tables.TryGetValue(tableName, out var rows);
                var result = rows?.Select(row => row.ToArray()).ToArray() ?? Array.Empty<string[]>();
                return new ReturnMessage(result, null, 0, call.LogicalCallContext, call);
            }

            public void SeedTable(string tableName, string[][] rows)
            {
                tables[tableName] = rows.Select(row => row?.ToArray() ?? Array.Empty<string>()).ToList();
            }

            public new object GetTransparentProxy()
            {
                return base.GetTransparentProxy();
            }
        }
    }
}
