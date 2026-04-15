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
        public void SaveBindingRoundTripsLayoutConfiguration()
        {
            var (store, adapter) = CreateStore();
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

            InvokeSaveBinding(store, binding);

            Assert.Equal("_Settings", adapter.WorksheetName);
            Assert.True(adapter.Visible);

            var loaded = InvokeLoadBinding(store, "Sheet1");

            Assert.Equal("performance", loaded.ProjectId);
            Assert.Equal("绩效项目", loaded.ProjectName);
            Assert.Equal(3, loaded.HeaderStartRow);
            Assert.Equal(2, loaded.HeaderRowCount);
            Assert.Equal(6, loaded.DataStartRow);
        }

        [Fact]
        public void SaveBindingPreservesOtherSheetBindings()
        {
            var (store, adapter) = CreateStore();
            adapter.SeedTable("SheetBindings", new[]
            {
                new[] { "Existing", "system-legacy", "legacy-project", "Legacy", "1", "2", "3" },
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
        public void SaveBindingRejectsBlankSheetName()
        {
            var (store, _) = CreateStore();
            var binding = new SheetBinding
            {
                SheetName = "  ",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
            };

            var error = Assert.Throws<TargetInvocationException>(() => InvokeSaveBinding(store, binding));
            Assert.IsType<ArgumentException>(error.InnerException);
        }

        [Fact]
        public void SaveFieldMappingsPreservesOtherSheetsAndUsesDynamicHeaders()
        {
            var (store, adapter) = CreateStore();
            var definition = new FieldMappingTableDefinition
            {
                SystemKey = "current-business-system",
                Columns = new[]
                {
                    new FieldMappingColumnDefinition
                    {
                        ColumnName = "HeaderId",
                        Role = FieldMappingSemanticRole.HeaderIdentity,
                    },
                    new FieldMappingColumnDefinition
                    {
                        ColumnName = "CurrentSingleDisplayName",
                        Role = FieldMappingSemanticRole.CurrentSingleHeaderText,
                    },
                },
            };

            adapter.SeedTable("SheetFieldMappings", new[]
            {
                new[] { "SheetA", "legacy_id", "旧列名" },
                new[] { "Sheet1", "old_sheet1_id", "旧负责人" },
            });

            InvokeSaveFieldMappings(
                store,
                "Sheet1",
                definition,
                new[]
                {
                    new SheetFieldMappingRow
                    {
                        SheetName = "Sheet1",
                        Values = new Dictionary<string, string>
                        {
                            ["HeaderId"] = "owner_name",
                            ["CurrentSingleDisplayName"] = "项目负责人",
                        },
                    },
                }
            );

            Assert.Equal("_Settings", adapter.WorksheetName);
            Assert.True(adapter.Visible);

            var loaded = InvokeLoadFieldMappings(store, "Sheet1", definition);
            var loadedRow = Assert.Single(loaded);
            Assert.Equal("owner_name", loadedRow.Values["HeaderId"]);
            Assert.Equal("项目负责人", loadedRow.Values["CurrentSingleDisplayName"]);

            var headers = adapter.ReadSeededHeaders("SheetFieldMappings");
            Assert.Equal(
                new[] { "SheetName", "HeaderId", "CurrentSingleDisplayName" },
                headers);

            var rawRows = adapter.ReadSeededTable("SheetFieldMappings");
            Assert.Contains(rawRows, row => row[0] == "SheetA" && row[1] == "legacy_id");
            Assert.DoesNotContain(rawRows, row => row[0] == "Sheet1" && row[1] == "old_sheet1_id");
        }

        [Fact]
        public void SaveFieldMappingsRejectsEmptyColumnNames()
        {
            var (store, _) = CreateStore();
            var definition = new FieldMappingTableDefinition
            {
                SystemKey = "current-business-system",
                Columns = new[]
                {
                    new FieldMappingColumnDefinition
                    {
                        ColumnName = " ",
                        Role = FieldMappingSemanticRole.HeaderIdentity,
                    },
                },
            };

            var error = Assert.Throws<TargetInvocationException>(() =>
                InvokeSaveFieldMappings(store, "Sheet1", definition, Array.Empty<SheetFieldMappingRow>()));
            Assert.IsType<ArgumentException>(error.InnerException);
        }

        [Fact]
        public void LoadFieldMappingsRejectsEmptyColumnNames()
        {
            var (store, _) = CreateStore();
            var definition = new FieldMappingTableDefinition
            {
                SystemKey = "current-business-system",
                Columns = new[]
                {
                    new FieldMappingColumnDefinition
                    {
                        ColumnName = "",
                        Role = FieldMappingSemanticRole.HeaderIdentity,
                    },
                },
            };

            var error = Assert.Throws<TargetInvocationException>(() =>
                InvokeLoadFieldMappings(store, "Sheet1", definition));
            Assert.IsType<ArgumentException>(error.InnerException);
        }

        [Fact]
        public void ClearFieldMappingsRemovesOnlyTargetSheetRowsAndPreservesHeaders()
        {
            var (store, adapter) = CreateStore();
            var definition = new FieldMappingTableDefinition
            {
                SystemKey = "current-business-system",
                Columns = new[]
                {
                    new FieldMappingColumnDefinition
                    {
                        ColumnName = "HeaderId",
                        Role = FieldMappingSemanticRole.HeaderIdentity,
                    },
                    new FieldMappingColumnDefinition
                    {
                        ColumnName = "CurrentSingleDisplayName",
                        Role = FieldMappingSemanticRole.CurrentSingleHeaderText,
                    },
                },
            };

            adapter.SeedTable("SheetFieldMappings", new[]
            {
                new[] { "SheetA", "legacy_id", "旧列名" },
                new[] { "Sheet1", "owner_name", "项目负责人" },
            });

            InvokeSaveFieldMappings(
                store,
                "Sheet1",
                definition,
                new[]
                {
                    new SheetFieldMappingRow
                    {
                        SheetName = "Sheet1",
                        Values = new Dictionary<string, string>
                        {
                            ["HeaderId"] = "owner_name",
                            ["CurrentSingleDisplayName"] = "项目负责人",
                        },
                    },
                });

            var headersBefore = adapter.ReadSeededHeaders("SheetFieldMappings");

            InvokeClearFieldMappings(store, "Sheet1");

            var rowsAfterClear = adapter.ReadSeededTable("SheetFieldMappings");
            Assert.Single(rowsAfterClear);
            Assert.Equal("SheetA", rowsAfterClear[0][0]);

            var headersAfter = adapter.ReadSeededHeaders("SheetFieldMappings");
            Assert.Equal(headersBefore, headersAfter);
            Assert.Equal("_Settings", adapter.WorksheetName);
            Assert.True(adapter.Visible);
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

        private static void InvokeSaveFieldMappings(
            object store,
            string sheetName,
            FieldMappingTableDefinition definition,
            IReadOnlyList<SheetFieldMappingRow> rows)
        {
            var method = store.GetType().GetMethod(
                "SaveFieldMappings",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            method.Invoke(store, new object[] { sheetName, definition, rows });
        }

        private static SheetFieldMappingRow[] InvokeLoadFieldMappings(
            object store,
            string sheetName,
            FieldMappingTableDefinition definition)
        {
            var method = store.GetType().GetMethod(
                "LoadFieldMappings",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            return (SheetFieldMappingRow[])method.Invoke(store, new object[] { sheetName, definition });
        }

        private static void InvokeClearFieldMappings(object store, string sheetName)
        {
            var method = store.GetType().GetMethod(
                "ClearFieldMappings",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            method.Invoke(store, new object[] { sheetName });
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
            private readonly Dictionary<string, string[]> headers =
                new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);

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
                var tableHeaders = (string[])call.InArgs[1];
                var rows = (string[][])call.InArgs[2];
                headers[tableName] = tableHeaders?.ToArray() ?? Array.Empty<string>();
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

            public string[][] ReadSeededTable(string tableName)
            {
                return tables.TryGetValue(tableName, out var rows)
                    ? rows.Select(row => row.ToArray()).ToArray()
                    : Array.Empty<string[]>();
            }

            public string[] ReadSeededHeaders(string tableName)
            {
                return headers.TryGetValue(tableName, out var tableHeaders)
                    ? tableHeaders.ToArray()
                    : Array.Empty<string>();
            }

            public new object GetTransparentProxy()
            {
                return base.GetTransparentProxy();
            }
        }
    }
}
