using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
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

            Assert.Equal("xISDP_Setting", adapter.WorksheetName);
            Assert.True(adapter.Visible);

            var loaded = InvokeLoadBinding(store, "Sheet1");

            Assert.Equal("performance", loaded.ProjectId);
            Assert.Equal("绩效项目", loaded.ProjectName);
            Assert.Equal(3, loaded.HeaderStartRow);
            Assert.Equal(2, loaded.HeaderRowCount);
            Assert.Equal(6, loaded.DataStartRow);

            var headers = adapter.ReadSeededHeaders("SheetBindings");
            Assert.Equal(
                new[]
                {
                    "SheetName",
                    "ProjectId",
                    "ProjectName",
                    "HeaderStartRow",
                    "HeaderRowCount",
                    "DataStartRow",
                    "SystemKey",
                },
                headers);

            var rawRow = Assert.Single(adapter.ReadSeededTable("SheetBindings"));
            Assert.Equal("performance", rawRow[1]);
            Assert.Equal("绩效项目", rawRow[2]);
            Assert.Equal("current-business-system", rawRow[6]);
        }

        [Fact]
        public void ImplementsWorksheetTemplateBindingStoreInterface()
        {
            var assembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var storeType = GetAddInType(assembly, "OfficeAgent.ExcelAddIn.Excel.WorksheetMetadataStore");

            Assert.Contains(
                typeof(IWorksheetTemplateBindingStore),
                storeType.GetInterfaces());
        }

        [Fact]
        public void SaveTemplateBindingRoundTripsAllTemplateColumns()
        {
            var (store, adapter) = CreateStore();
            var appliedAt = new DateTime(2026, 4, 22, 9, 30, 15, DateTimeKind.Utc);
            var binding = new SheetTemplateBinding
            {
                SheetName = "Sheet1",
                TemplateId = "template-001",
                TemplateName = "Quarterly Report",
                TemplateRevision = 5,
                TemplateOrigin = "local",
                AppliedFingerprint = "fingerprint-123",
                TemplateLastAppliedAt = appliedAt,
                DerivedFromTemplateId = "template-base",
                DerivedFromTemplateRevision = 2,
            };

            InvokeSaveTemplateBinding(store, binding);

            Assert.Equal("xISDP_Setting", adapter.WorksheetName);
            Assert.True(adapter.Visible);

            var loaded = InvokeLoadTemplateBinding(store, "Sheet1");

            Assert.Equal("template-001", loaded.TemplateId);
            Assert.Equal("Quarterly Report", loaded.TemplateName);
            Assert.Equal(5, loaded.TemplateRevision);
            Assert.Equal("local", loaded.TemplateOrigin);
            Assert.Equal("fingerprint-123", loaded.AppliedFingerprint);
            Assert.Equal(appliedAt, loaded.TemplateLastAppliedAt);
            Assert.Equal("template-base", loaded.DerivedFromTemplateId);
            Assert.Equal(2, loaded.DerivedFromTemplateRevision);

            var headers = adapter.ReadSeededHeaders("TemplateBindings");
            Assert.Equal(
                new[]
                {
                    "SheetName",
                    "TemplateName",
                    "TemplateRevision",
                    "TemplateOrigin",
                    "TemplateId",
                    "AppliedFingerprint",
                    "TemplateLastAppliedAt",
                    "DerivedFromTemplateId",
                    "DerivedFromTemplateRevision",
                },
                headers);

            var rawRow = Assert.Single(adapter.ReadSeededTable("TemplateBindings"));
            Assert.Equal("Quarterly Report", rawRow[1]);
            Assert.Equal("5", rawRow[2]);
            Assert.Equal("local", rawRow[3]);
            Assert.Equal("template-001", rawRow[4]);
        }

        [Fact]
        public void ClearTemplateBindingRemovesOnlyTargetSheetRows()
        {
            var (store, adapter) = CreateStore();
            adapter.SeedTable("TemplateBindings", new[]
            {
                new[] { "SheetA", "Template A", "1", "local", "template-a", "fp-a", "2026-04-22T10:00:00.0000000Z", "base-a", "1" },
                new[] { "SheetB", "Template B", "2", "local", "template-b", "fp-b", "2026-04-22T11:00:00.0000000Z", "base-b", "2" },
            });

            InvokeClearTemplateBinding(store, "SheetA");

            var rowsAfterClear = adapter.ReadSeededTable("TemplateBindings");
            var remaining = Assert.Single(rowsAfterClear);
            Assert.Equal("SheetB", remaining[0]);
            Assert.Equal("Template B", remaining[1]);
            Assert.Equal("template-b", remaining[4]);
        }

        [Fact]
        public void InvalidateCacheForcesTemplateBindingsToReloadFromAdapter()
        {
            var (store, adapter) = CreateStore();
            adapter.SeedTable("TemplateBindings", new[]
            {
                new[] { "Sheet1", "Template A", "1", "local", "template-a", "fp-a", "2026-04-22T10:00:00.0000000Z", string.Empty, string.Empty },
            });

            var first = InvokeLoadTemplateBinding(store, "Sheet1");
            Assert.Equal("template-a", first.TemplateId);
            Assert.Equal(1, adapter.ReadTableCallCount);

            adapter.SeedTable("TemplateBindings", new[]
            {
                new[] { "Sheet1", "Template B", "3", "local", "template-b", "fp-b", "2026-04-22T11:00:00.0000000Z", string.Empty, string.Empty },
            });

            InvokeInvalidateCache(store);
            var refreshed = InvokeLoadTemplateBinding(store, "Sheet1");

            Assert.Equal("template-b", refreshed.TemplateId);
            Assert.Equal(2, adapter.ReadTableCallCount);
        }

        [Fact]
        public void SaveBindingPreservesOtherSheetBindings()
        {
            var (store, adapter) = CreateStore();
            adapter.SeedTable("SheetBindings", new[]
            {
                new[] { "Existing", "legacy-project", "Legacy", "1", "2", "3", "system-legacy" },
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
        public void LoadBindingDoesNotCreateSettingsWorksheetWhenMetadataIsMissing()
        {
            var (store, adapter) = CreateStore();

            var error = Assert.Throws<TargetInvocationException>(() => InvokeLoadBinding(store, "Sheet1"));

            Assert.IsType<InvalidOperationException>(error.InnerException);
            Assert.Equal(0, adapter.EnsureWorksheetCallCount);
            Assert.Null(adapter.WorksheetName);
        }

        [Fact]
        public void LoadBindingUsesCachedRowsUntilBindingsChange()
        {
            var (store, adapter) = CreateStore();
            adapter.SeedTable("SheetBindings", new[]
            {
                new[] { "Sheet1", "performance", "绩效项目", "3", "2", "6", "current-business-system" },
            });

            var first = InvokeLoadBinding(store, "Sheet1");
            var second = InvokeLoadBinding(store, "Sheet1");

            Assert.Equal("performance", first.ProjectId);
            Assert.Equal("performance", second.ProjectId);
            Assert.Equal(1, adapter.ReadTableCallCount);

            InvokeSaveBinding(store, new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "refreshed-project",
                ProjectName = "刷新项目",
                HeaderStartRow = 4,
                HeaderRowCount = 1,
                DataStartRow = 8,
            });

            var refreshed = InvokeLoadBinding(store, "Sheet1");

            Assert.Equal("refreshed-project", refreshed.ProjectId);
            Assert.Equal(1, adapter.ReadTableCallCount);
        }

        [Fact]
        public void InvalidateCacheForcesBindingRowsToReloadFromAdapter()
        {
            var (store, adapter) = CreateStore();
            adapter.SeedTable("SheetBindings", new[]
            {
                new[] { "Sheet1", "performance", "绩效项目", "3", "2", "6", "current-business-system" },
            });

            var first = InvokeLoadBinding(store, "Sheet1");
            Assert.Equal("performance", first.ProjectId);
            Assert.Equal(1, adapter.ReadTableCallCount);

            adapter.SeedTable("SheetBindings", new[]
            {
                new[] { "Sheet1", "updated-project", "新项目", "4", "1", "8", "current-business-system" },
            });

            InvokeInvalidateCache(store);
            var refreshed = InvokeLoadBinding(store, "Sheet1");

            Assert.Equal("updated-project", refreshed.ProjectId);
            Assert.Equal(2, adapter.ReadTableCallCount);
        }

        [Fact]
        public void LoadBindingReloadsRowsWhenWorkbookScopeChanges()
        {
            var (store, adapter) = CreateStore();
            adapter.SwitchWorkbook("WorkbookA");
            adapter.SeedTable("SheetBindings", new[]
            {
                new[] { "Sheet1", "project-a", "项目A", "3", "2", "6", "current-business-system" },
            });

            var first = InvokeLoadBinding(store, "Sheet1");

            Assert.Equal("project-a", first.ProjectId);
            Assert.Equal(1, adapter.ReadTableCallCount);

            adapter.SwitchWorkbook("WorkbookB");
            adapter.SeedTable("SheetBindings", new[]
            {
                new[] { "Sheet1", "project-b", "项目B", "4", "1", "8", "current-business-system" },
            });

            var second = InvokeLoadBinding(store, "Sheet1");

            Assert.Equal("project-b", second.ProjectId);
            Assert.Equal(2, adapter.ReadTableCallCount);
        }

        [Fact]
        public void SaveBindingKeepsWorkbookRowsIsolatedWhenWorkbookScopeChanges()
        {
            var (store, adapter) = CreateStore();
            adapter.SwitchWorkbook("WorkbookA");
            adapter.SeedTable("SheetBindings", new[]
            {
                new[] { "SheetA", "project-a", "项目A", "1", "2", "3", "current-business-system" },
            });

            var original = InvokeLoadBinding(store, "SheetA");
            Assert.Equal("project-a", original.ProjectId);

            adapter.SwitchWorkbook("WorkbookB");
            adapter.SeedTable("SheetBindings", new[]
            {
                new[] { "SheetB", "project-b", "项目B", "2", "1", "4", "current-business-system" },
            });

            InvokeSaveBinding(store, new SheetBinding
            {
                SheetName = "SheetC",
                SystemKey = "current-business-system",
                ProjectId = "project-c",
                ProjectName = "项目C",
                HeaderStartRow = 5,
                HeaderRowCount = 2,
                DataStartRow = 8,
            });

            var workbookBRows = adapter.ReadSeededTable("SheetBindings");

            Assert.Contains(workbookBRows, row => row[0] == "SheetB" && row[1] == "project-b");
            Assert.Contains(workbookBRows, row => row[0] == "SheetC" && row[1] == "project-c");
            Assert.DoesNotContain(workbookBRows, row => row[0] == "SheetA");
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
                        ColumnName = "HeaderType",
                        Role = FieldMappingSemanticRole.HeaderType,
                    },
                    new FieldMappingColumnDefinition
                    {
                        ColumnName = "ISDP L1",
                        Role = FieldMappingSemanticRole.DefaultSingleHeaderText,
                        RoleKey = "DefaultL1",
                    },
                    new FieldMappingColumnDefinition
                    {
                        ColumnName = "Excel L1",
                        Role = FieldMappingSemanticRole.CurrentSingleHeaderText,
                        RoleKey = "CurrentL1",
                    },
                    new FieldMappingColumnDefinition
                    {
                        ColumnName = "ISDP L1",
                        Role = FieldMappingSemanticRole.DefaultParentHeaderText,
                        RoleKey = "DefaultL1",
                    },
                    new FieldMappingColumnDefinition
                    {
                        ColumnName = "Excel L1",
                        Role = FieldMappingSemanticRole.CurrentParentHeaderText,
                        RoleKey = "CurrentL1",
                    },
                    new FieldMappingColumnDefinition
                    {
                        ColumnName = "ISDP L2",
                        Role = FieldMappingSemanticRole.DefaultChildHeaderText,
                        RoleKey = "DefaultL2",
                    },
                    new FieldMappingColumnDefinition
                    {
                        ColumnName = "Excel L2",
                        Role = FieldMappingSemanticRole.CurrentChildHeaderText,
                        RoleKey = "CurrentL2",
                    },
                    new FieldMappingColumnDefinition
                    {
                        ColumnName = "HeaderId",
                        Role = FieldMappingSemanticRole.HeaderIdentity,
                    },
                },
            };

            adapter.SeedTable("SheetFieldMappings", new[]
            {
                new[] { "SheetA", "single", "旧L1", "当前旧L1", string.Empty, string.Empty, "legacy_id" },
                new[] { "Sheet1", "single", "旧负责人", "当前旧负责人", string.Empty, string.Empty, "old_sheet1_id" },
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
                            ["HeaderType"] = "single",
                            ["DefaultL1"] = "负责人",
                            ["CurrentL1"] = "项目负责人",
                            ["DefaultL2"] = string.Empty,
                            ["CurrentL2"] = string.Empty,
                            ["HeaderId"] = "owner_name",
                        },
                    },
                }
            );

            Assert.Equal("xISDP_Setting", adapter.WorksheetName);
            Assert.True(adapter.Visible);

            var loaded = InvokeLoadFieldMappings(store, "Sheet1", definition);
            var loadedRow = Assert.Single(loaded);
            Assert.Equal("single", loadedRow.Values["HeaderType"]);
            Assert.Equal("负责人", loadedRow.Values["DefaultL1"]);
            Assert.Equal("项目负责人", loadedRow.Values["CurrentL1"]);
            Assert.Equal("owner_name", loadedRow.Values["HeaderId"]);

            var headers = adapter.ReadSeededHeaders("SheetFieldMappings");
            Assert.Equal(
                new[] { "SheetName", "HeaderType", "ISDP L1", "Excel L1", "ISDP L2", "Excel L2", "HeaderId" },
                headers);

            var rawRows = adapter.ReadSeededTable("SheetFieldMappings");
            Assert.Contains(rawRows, row => row[0] == "SheetA" && row[6] == "legacy_id");
            Assert.DoesNotContain(rawRows, row => row[0] == "Sheet1" && row[6] == "old_sheet1_id");
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
        public void LoadFieldMappingsDoesNotCreateSettingsWorksheetWhenMetadataIsMissing()
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
                },
            };

            var rows = InvokeLoadFieldMappings(store, "Sheet1", definition);

            Assert.Empty(rows);
            Assert.Equal(0, adapter.EnsureWorksheetCallCount);
            Assert.Null(adapter.WorksheetName);
        }

        [Fact]
        public void LoadFieldMappingsUsesCachedRowsUntilMappingsChange()
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
                new[] { "Sheet1", "owner_name", "项目负责人" },
            });

            var first = InvokeLoadFieldMappings(store, "Sheet1", definition);
            var second = InvokeLoadFieldMappings(store, "Sheet1", definition);

            Assert.Single(first);
            Assert.Single(second);
            Assert.Equal(1, adapter.ReadTableCallCount);

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
                            ["HeaderId"] = "row_id",
                            ["CurrentSingleDisplayName"] = "ID",
                        },
                    },
                });

            var refreshed = InvokeLoadFieldMappings(store, "Sheet1", definition);

            Assert.Single(refreshed);
            Assert.Equal("row_id", refreshed[0].Values["HeaderId"]);
            Assert.Equal(1, adapter.ReadTableCallCount);
        }

        [Fact]
        public void InvalidateCacheForcesFieldMappingsToReloadFromAdapter()
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
                new[] { "Sheet1", "owner_name", "项目负责人" },
            });

            var first = InvokeLoadFieldMappings(store, "Sheet1", definition);
            Assert.Single(first);
            Assert.Equal(1, adapter.ReadTableCallCount);

            adapter.SeedTable("SheetFieldMappings", new[]
            {
                new[] { "Sheet1", "row_id", "ID" },
            });

            InvokeInvalidateCache(store);
            var refreshed = InvokeLoadFieldMappings(store, "Sheet1", definition);

            Assert.Single(refreshed);
            Assert.Equal("row_id", refreshed[0].Values["HeaderId"]);
            Assert.Equal(2, adapter.ReadTableCallCount);
        }

        [Fact]
        public void LoadFieldMappingsReloadsRowsWhenWorkbookScopeChanges()
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

            adapter.SwitchWorkbook("WorkbookA");
            adapter.SeedTable("SheetFieldMappings", new[]
            {
                new[] { "Sheet1", "owner_name", "项目负责人A" },
            });

            var first = InvokeLoadFieldMappings(store, "Sheet1", definition);

            Assert.Single(first);
            Assert.Equal("项目负责人A", first[0].Values["CurrentSingleDisplayName"]);
            Assert.Equal(1, adapter.ReadTableCallCount);

            adapter.SwitchWorkbook("WorkbookB");
            adapter.SeedTable("SheetFieldMappings", new[]
            {
                new[] { "Sheet1", "row_id", "项目负责人B" },
            });

            var second = InvokeLoadFieldMappings(store, "Sheet1", definition);

            Assert.Single(second);
            Assert.Equal("项目负责人B", second[0].Values["CurrentSingleDisplayName"]);
            Assert.Equal(2, adapter.ReadTableCallCount);
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
            Assert.Equal("xISDP_Setting", adapter.WorksheetName);
            Assert.True(adapter.Visible);
        }

        [Fact]
        public void ClearFieldMappingsDoesNotCreateSettingsWorksheetWhenMetadataIsMissing()
        {
            var (store, adapter) = CreateStore();

            InvokeClearFieldMappings(store, "Sheet1");

            Assert.Equal(0, adapter.EnsureWorksheetCallCount);
            Assert.Null(adapter.WorksheetName);
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

        private static void InvokeSaveTemplateBinding(object store, SheetTemplateBinding binding)
        {
            var method = store.GetType().GetMethod(
                "SaveTemplateBinding",
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

        private static SheetTemplateBinding InvokeLoadTemplateBinding(object store, string sheetName)
        {
            var method = store.GetType().GetMethod(
                "LoadTemplateBinding",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            return (SheetTemplateBinding)method.Invoke(store, new object[] { sheetName });
        }

        private static void InvokeClearTemplateBinding(object store, string sheetName)
        {
            var method = store.GetType().GetMethod(
                "ClearTemplateBinding",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            method.Invoke(store, new object[] { sheetName });
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

        private static void InvokeInvalidateCache(object store)
        {
            var method = store.GetType().GetMethod(
                "InvalidateCache",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            method.Invoke(store, null);
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
            private readonly Dictionary<string, Dictionary<string, List<string[]>>> tablesByWorkbook =
                new Dictionary<string, Dictionary<string, List<string[]>>>(StringComparer.OrdinalIgnoreCase);
            private readonly Dictionary<string, Dictionary<string, string[]>> headersByWorkbook =
                new Dictionary<string, Dictionary<string, string[]>>(StringComparer.OrdinalIgnoreCase);

            public int ReadTableCallCount { get; private set; }
            public string WorksheetName { get; private set; }
            public bool Visible { get; private set; }
            public int EnsureWorksheetCallCount { get; private set; }
            public string WorkbookScopeKey { get; private set; } = "Workbook1";

            public FakeWorksheetMetadataAdapter(Type adapterInterface)
                : base(adapterInterface)
            {
            }

            public override IMessage Invoke(IMessage msg)
            {
                var call = (IMethodCallMessage)msg;

                return call.MethodName switch
                {
                    "GetWorkbookScopeKey" => HandleGetWorkbookScopeKey(call),
                    "EnsureWorksheet" => HandleEnsureWorksheet(call),
                    "WriteTable" => HandleWriteTable(call),
                    "ReadTable" => HandleReadTable(call),
                    _ => throw new NotSupportedException(call.MethodName),
                };
            }

            private IMessage HandleGetWorkbookScopeKey(IMethodCallMessage call)
            {
                return new ReturnMessage(WorkbookScopeKey, null, 0, call.LogicalCallContext, call);
            }

            private IMessage HandleEnsureWorksheet(IMethodCallMessage call)
            {
                EnsureWorksheetCallCount++;
                WorksheetName = (string)call.InArgs[0];
                Visible = (bool)call.InArgs[1];
                return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
            }

            private IMessage HandleWriteTable(IMethodCallMessage call)
            {
                var tableName = (string)call.InArgs[0];
                var tableHeaders = (string[])call.InArgs[1];
                var rows = (string[][])call.InArgs[2];
                var headers = GetCurrentWorkbookHeaders();
                var tables = GetCurrentWorkbookTables();
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
                ReadTableCallCount++;
                var tableName = (string)call.InArgs[0];
                var tables = GetCurrentWorkbookTables();
                tables.TryGetValue(tableName, out var rows);
                var result = rows?.Select(row => row.ToArray()).ToArray() ?? Array.Empty<string[]>();
                return new ReturnMessage(result, null, 0, call.LogicalCallContext, call);
            }

            public void SeedTable(string tableName, string[][] rows)
            {
                var tables = GetCurrentWorkbookTables();
                tables[tableName] = rows.Select(row => row?.ToArray() ?? Array.Empty<string>()).ToList();
            }

            public string[][] ReadSeededTable(string tableName)
            {
                var tables = GetCurrentWorkbookTables();
                return tables.TryGetValue(tableName, out var rows)
                    ? rows.Select(row => row.ToArray()).ToArray()
                    : Array.Empty<string[]>();
            }

            public string[] ReadSeededHeaders(string tableName)
            {
                var headers = GetCurrentWorkbookHeaders();
                return headers.TryGetValue(tableName, out var tableHeaders)
                    ? tableHeaders.ToArray()
                    : Array.Empty<string>();
            }

            public void SwitchWorkbook(string workbookScopeKey)
            {
                WorkbookScopeKey = workbookScopeKey ?? string.Empty;
                GetCurrentWorkbookTables();
                GetCurrentWorkbookHeaders();
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
    }
}
