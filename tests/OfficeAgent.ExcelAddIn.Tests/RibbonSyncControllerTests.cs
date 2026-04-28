using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using OfficeAgent.Core;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
using OfficeAgent.Core.Sync;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class RibbonSyncControllerTests
    {
        [Fact]
        public void NewControllerDefaultsToSelectProjectDisplayWhenNoBinding()
        {
            var controller = CreateController(new FakeSystemConnector(), new FakeWorksheetMetadataStore(), new FakeDialogService(), () => "Sheet1");

            Assert.Equal("先选择项目", ReadActiveProjectDisplayName(controller));
        }

        [Fact]
        public void SelectProjectShowsLayoutDialogAndSavesConfirmedBindingWithoutAutoInitialize()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var dialogService = new FakeDialogService
            {
                NextProjectLayoutBinding = new SheetBinding
                {
                    SheetName = "Sheet1",
                    SystemKey = "current-business-system",
                    ProjectId = "performance",
                    ProjectName = "绩效项目",
                    HeaderStartRow = 4,
                    HeaderRowCount = 1,
                    DataStartRow = 5,
                },
            };
            var controller = CreateController(connector, metadataStore, dialogService, () => "Sheet1");
            var option = new ProjectOption
            {
                SystemKey = "current-business-system",
                ProjectId = "performance",
                DisplayName = "绩效项目",
            };

            InvokeSelectProject(controller, option);

            Assert.Single(dialogService.ProjectLayoutPrompts);
            Assert.Equal(1, dialogService.ProjectLayoutPrompts[0].HeaderStartRow);
            Assert.Equal(2, dialogService.ProjectLayoutPrompts[0].HeaderRowCount);
            Assert.Equal(3, dialogService.ProjectLayoutPrompts[0].DataStartRow);
            Assert.NotNull(metadataStore.LastSavedBinding);
            Assert.Equal("Sheet1", metadataStore.LastSavedBinding.SheetName);
            Assert.Equal("performance", metadataStore.LastSavedBinding.ProjectId);
            Assert.Equal(4, metadataStore.LastSavedBinding.HeaderStartRow);
            Assert.Equal(1, metadataStore.LastSavedBinding.HeaderRowCount);
            Assert.Equal(5, metadataStore.LastSavedBinding.DataStartRow);
            Assert.Equal("绩效项目", ReadActiveProjectDisplayName(controller));
            Assert.Empty(metadataStore.LastSavedFieldMappings);
            Assert.Null(connector.LastBuildFieldMappingSeedProjectId);
            Assert.Empty(dialogService.WarningMessages);
        }

        [Fact]
        public void SelectProjectUsesExistingLayoutAsDialogDefaultsWhenSwitchingProject()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var dialogService = new FakeDialogService
            {
                NextProjectLayoutBinding = new SheetBinding
                {
                    SheetName = "Sheet1",
                    SystemKey = "current-business-system",
                    ProjectId = "new-project",
                    ProjectName = "新项目",
                    HeaderStartRow = 5,
                    HeaderRowCount = 2,
                    DataStartRow = 7,
                },
            };
            metadataStore.Bindings["Sheet1"] = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "old-project",
                ProjectName = "旧项目",
                HeaderStartRow = 5,
                HeaderRowCount = 2,
                DataStartRow = 7,
            };

            var controller = CreateController(connector, metadataStore, dialogService, () => "Sheet1");
            InvokeSelectProject(controller, new ProjectOption
            {
                SystemKey = "current-business-system",
                ProjectId = "new-project",
                DisplayName = "新项目",
            });

            Assert.Single(dialogService.ProjectLayoutPrompts);
            Assert.Equal("Sheet1", dialogService.ProjectLayoutPrompts[0].SheetName);
            Assert.Equal("current-business-system", dialogService.ProjectLayoutPrompts[0].SystemKey);
            Assert.Equal("new-project", dialogService.ProjectLayoutPrompts[0].ProjectId);
            Assert.Equal("新项目", dialogService.ProjectLayoutPrompts[0].ProjectName);
            Assert.Equal(5, dialogService.ProjectLayoutPrompts[0].HeaderStartRow);
            Assert.Equal(2, dialogService.ProjectLayoutPrompts[0].HeaderRowCount);
            Assert.Equal(7, dialogService.ProjectLayoutPrompts[0].DataStartRow);
        }

        [Fact]
        public void SelectProjectDoesNotPromptOrSaveWhenSameProjectIsReselected()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var dialogService = new FakeDialogService();
            metadataStore.Bindings["Sheet1"] = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "old-project",
                ProjectName = "旧项目",
                HeaderStartRow = 5,
                HeaderRowCount = 2,
                DataStartRow = 7,
            };

            var controller = CreateController(connector, metadataStore, dialogService, () => "Sheet1");
            InvokeRefresh(controller);

            InvokeSelectProject(controller, new ProjectOption
            {
                SystemKey = "current-business-system",
                ProjectId = "old-project",
                DisplayName = "旧项目",
            });

            Assert.Empty(dialogService.ProjectLayoutPrompts);
            Assert.Null(metadataStore.LastSavedBinding);
            Assert.Equal("old-project", ReadActiveProjectId(controller));
            Assert.Equal("旧项目", ReadActiveProjectDisplayName(controller));
        }

        [Fact]
        public void SelectProjectCancelKeepsExistingBindingAndActiveProjectState()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var dialogService = new FakeDialogService
            {
                NextProjectLayoutBinding = null,
            };
            metadataStore.Bindings["Sheet1"] = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "old-project",
                ProjectName = "旧项目",
                HeaderStartRow = 5,
                HeaderRowCount = 2,
                DataStartRow = 7,
            };

            var controller = CreateController(connector, metadataStore, dialogService, () => "Sheet1");
            InvokeRefresh(controller);

            InvokeSelectProject(controller, new ProjectOption
            {
                SystemKey = "current-business-system",
                ProjectId = "new-project",
                DisplayName = "新项目",
            });

            Assert.Single(dialogService.ProjectLayoutPrompts);
            Assert.Null(metadataStore.LastSavedBinding);
            Assert.Equal("old-project", ReadActiveProjectId(controller));
            Assert.Equal("旧项目", ReadActiveProjectDisplayName(controller));
        }

        [Fact]
        public void SelectProjectClearsFieldMappingsWhenSwitchingToDifferentProject()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var dialogService = new FakeDialogService
            {
                NextProjectLayoutBinding = new SheetBinding
                {
                    SheetName = "Sheet1",
                    SystemKey = "current-business-system",
                    ProjectId = "new-project",
                    ProjectName = "新项目",
                    HeaderStartRow = 1,
                    HeaderRowCount = 2,
                    DataStartRow = 3,
                },
            };
            metadataStore.Bindings["Sheet1"] = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "old-project",
                ProjectName = "旧项目",
                HeaderStartRow = 1,
                HeaderRowCount = 2,
                DataStartRow = 3,
            };
            metadataStore.FieldMappings["Sheet1"] = new[]
            {
                new SheetFieldMappingRow
                {
                    SheetName = "Sheet1",
                    Values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["ApiFieldKey"] = "row_id",
                    },
                },
            };

            var controller = CreateController(connector, metadataStore, dialogService, () => "Sheet1");

            InvokeSelectProject(controller, new ProjectOption
            {
                SystemKey = "current-business-system",
                ProjectId = "new-project",
                DisplayName = "新项目",
            });

            Assert.False(metadataStore.FieldMappings.ContainsKey("Sheet1"));
            Assert.NotNull(metadataStore.LastSavedBinding);
            Assert.Equal("new-project", metadataStore.LastSavedBinding.ProjectId);
            Assert.Equal("新项目", metadataStore.LastSavedBinding.ProjectName);
        }

        [Fact]
        public void ExecuteInitializeCurrentSheetPreservesSavedLayoutAndReportsSuccess()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var dialogService = new FakeDialogService();
            metadataStore.Bindings["Sheet1"] = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 5,
                HeaderRowCount = 2,
                DataStartRow = 9,
            };

            var controller = CreateController(connector, metadataStore, dialogService, () => "Sheet1");
            InvokeRefresh(controller);

            InvokeExecuteInitializeCurrentSheet(controller);

            Assert.Equal("performance", connector.LastBuildFieldMappingSeedProjectId);
            Assert.Equal(5, metadataStore.LastSavedBinding.HeaderStartRow);
            Assert.Equal(2, metadataStore.LastSavedBinding.HeaderRowCount);
            Assert.Equal(9, metadataStore.LastSavedBinding.DataStartRow);
            Assert.Contains(dialogService.InfoMessages, message => message.Contains("初始化当前表完成"));
        }

        [Fact]
        public void ExecuteInitializeCurrentSheetPromptsLoginWhenAuthenticationIsRequired()
        {
            var connector = new FakeSystemConnector
            {
                BuildFieldMappingSeedException = new AuthenticationRequiredException("当前未登录，请先登录"),
            };
            var metadataStore = new FakeWorksheetMetadataStore();
            var dialogService = new FakeDialogService
            {
                AuthenticationRequiredResult = true,
            };
            var loginTriggered = false;
            metadataStore.Bindings["Sheet1"] = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 5,
                HeaderRowCount = 2,
                DataStartRow = 9,
            };

            var controller = CreateController(
                connector,
                metadataStore,
                dialogService,
                () => "Sheet1",
                () =>
                {
                    loginTriggered = true;
                });
            InvokeRefresh(controller);

            InvokeExecuteInitializeCurrentSheet(controller);

            Assert.Single(dialogService.AuthenticationRequiredMessages);
            Assert.Equal("当前未登录，请先登录", dialogService.AuthenticationRequiredMessages[0]);
            Assert.True(loginTriggered);
            Assert.Empty(dialogService.ErrorMessages);
        }

        [Fact]
        public void RefreshActiveProjectFromSheetMetadataLoadsBindingForCurrentSheet()
        {
            var metadataStore = new FakeWorksheetMetadataStore();
            metadataStore.Bindings["SheetWithBinding"] = new SheetBinding
            {
                SheetName = "SheetWithBinding",
                SystemKey = "current-business-system",
                ProjectId = "project-2",
                ProjectName = "项目二",
            };

            var controller = CreateController(new FakeSystemConnector(), metadataStore, new FakeDialogService(), () => "SheetWithBinding");

            InvokeRefresh(controller);

            Assert.Equal("项目二", ReadActiveProjectDisplayName(controller));
            Assert.Equal("project-2", ReadActiveProjectId(controller));
        }

        [Fact]
        public void RefreshActiveProjectFromSheetMetadataSkipsReloadWhenActiveSheetDidNotChange()
        {
            var metadataStore = new FakeWorksheetMetadataStore();
            metadataStore.Bindings["SheetWithBinding"] = new SheetBinding
            {
                SheetName = "SheetWithBinding",
                SystemKey = "current-business-system",
                ProjectId = "project-2",
                ProjectName = "项目二",
            };

            var controller = CreateController(new FakeSystemConnector(), metadataStore, new FakeDialogService(), () => "SheetWithBinding");

            InvokeRefresh(controller);
            InvokeRefresh(controller);

            Assert.Equal(1, metadataStore.LoadBindingCallCount);
            Assert.Equal("项目二", ReadActiveProjectDisplayName(controller));
        }

        [Fact]
        public void InvalidatingRefreshStateForcesReloadForSameActiveSheet()
        {
            var metadataStore = new FakeWorksheetMetadataStore();
            metadataStore.Bindings["SheetWithBinding"] = new SheetBinding
            {
                SheetName = "SheetWithBinding",
                SystemKey = "current-business-system",
                ProjectId = "project-2",
                ProjectName = "项目二",
            };

            var controller = CreateController(new FakeSystemConnector(), metadataStore, new FakeDialogService(), () => "SheetWithBinding");

            InvokeRefresh(controller);
            InvokeInvalidateRefreshState(controller);
            InvokeRefresh(controller);

            Assert.Equal(2, metadataStore.LoadBindingCallCount);
        }

        [Fact]
        public void RefreshProjectFromExplicitSheetNameUsesActivatedSheetEvenWhenActiveSheetProviderIsStale()
        {
            var metadataStore = new FakeWorksheetMetadataStore();
            metadataStore.Bindings["SheetA"] = new SheetBinding
            {
                SheetName = "SheetA",
                SystemKey = "current-business-system",
                ProjectId = "project-a",
                ProjectName = "项目A",
            };
            metadataStore.Bindings["SheetB"] = new SheetBinding
            {
                SheetName = "SheetB",
                SystemKey = "current-business-system",
                ProjectId = "project-b",
                ProjectName = "项目B",
            };

            var controller = CreateController(new FakeSystemConnector(), metadataStore, new FakeDialogService(), () => "SheetA");

            InvokeRefresh(controller);
            InvokeRefreshForSheet(controller, "SheetB");

            Assert.Equal("project-b", ReadActiveProjectId(controller));
            Assert.Equal("项目B", ReadActiveProjectDisplayName(controller));
        }

        [Fact]
        public void RefreshActiveProjectFromSheetMetadataFallsBackToDefaultWhenBindingMissing()
        {
            var metadataStore = new FakeWorksheetMetadataStore();
            var controller = CreateController(new FakeSystemConnector(), metadataStore, new FakeDialogService(), () => "SheetWithoutBinding");

            InvokeRefresh(controller);

            Assert.Equal("先选择项目", ReadActiveProjectDisplayName(controller));
            Assert.Equal(string.Empty, ReadActiveProjectId(controller));
        }

        [Fact]
        public void RefreshActiveProjectFromSettingsSheetFallsBackToDefaultWhenSettingsSheetHasNoBinding()
        {
            var metadataStore = new FakeWorksheetMetadataStore();
            metadataStore.Bindings["Sheet1"] = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
            };

            var controller = CreateController(new FakeSystemConnector(), metadataStore, new FakeDialogService(), () => "ISDP_Setting");

            InvokeRefresh(controller);

            Assert.Equal(string.Empty, ReadActiveProjectId(controller));
            Assert.Equal("先选择项目", ReadActiveProjectDisplayName(controller));
        }

        [Fact]
        public void RefreshProjectFromSettingsSheetClearsPreviousBusinessProjectStateWhenSettingsSheetHasNoBinding()
        {
            var metadataStore = new FakeWorksheetMetadataStore();
            metadataStore.Bindings["SheetA"] = new SheetBinding
            {
                SheetName = "SheetA",
                SystemKey = "current-business-system",
                ProjectId = "project-a",
                ProjectName = "项目A",
            };
            metadataStore.Bindings["SheetB"] = new SheetBinding
            {
                SheetName = "SheetB",
                SystemKey = "current-business-system",
                ProjectId = "project-b",
                ProjectName = "项目B",
            };

            var controller = CreateController(new FakeSystemConnector(), metadataStore, new FakeDialogService(), () => "SheetA");

            InvokeRefresh(controller);
            InvokeRefreshForSheet(controller, "ISDP_Setting");

            Assert.Equal(string.Empty, ReadActiveProjectId(controller));
            Assert.Equal("先选择项目", ReadActiveProjectDisplayName(controller));
        }

        [Fact]
        public void RefreshActiveProjectFromSettingsSheetLoadsBindingWhenSettingsSheetIsExplicitlyBound()
        {
            var metadataStore = new FakeWorksheetMetadataStore();
            metadataStore.Bindings["ISDP_Setting"] = new SheetBinding
            {
                SheetName = "ISDP_Setting",
                SystemKey = "current-business-system",
                ProjectId = "settings-project",
                ProjectName = "设置页项目",
            };

            var controller = CreateController(new FakeSystemConnector(), metadataStore, new FakeDialogService(), () => "ISDP_Setting");

            InvokeRefresh(controller);

            Assert.Equal("settings-project", ReadActiveProjectId(controller));
            Assert.Equal("设置页项目", ReadActiveProjectDisplayName(controller));
        }

        [Fact]
        public void RefreshActiveProjectFromSheetMetadataWithBlankProjectNameFallsBackToProjectIdLabel()
        {
            var metadataStore = new FakeWorksheetMetadataStore();
            metadataStore.Bindings["SheetWithBinding"] = new SheetBinding
            {
                SheetName = "SheetWithBinding",
                SystemKey = "current-business-system",
                ProjectId = "project-2",
                ProjectName = "   ",
            };

            var controller = CreateController(new FakeSystemConnector(), metadataStore, new FakeDialogService(), () => "SheetWithBinding");
            InvokeRefresh(controller);

            Assert.Equal(string.Empty, ReadActiveProjectDisplayName(controller));
            Assert.Equal("project-2", ReadActiveProjectId(controller));
            Assert.Equal(
                "project-2",
                InvokeFormatProjectDropDownLabel(ReadActiveProjectId(controller), ReadActiveProjectDisplayName(controller)));
        }

        [Fact]
        public void BuildUploadPreviewInfoMessageIncludesSkippedReasonsWhenNothingWillUpload()
        {
            var preview = new SyncOperationPreview
            {
                Summary = "部分上传将上传 0 个单元格，跳过 1 个单元格。",
                Details = new[] { "row-1 / status: 已跳过，单据已归档，禁止上传" },
                Changes = Array.Empty<CellChange>(),
                SkippedChanges = new[]
                {
                    new SkippedCellChange
                    {
                        Change = new CellChange { RowId = "row-1", ApiFieldKey = "status" },
                        Reason = "单据已归档，禁止上传",
                    },
                },
            };

            var message = InvokeBuildUploadPreviewInfoMessage("部分上传", preview);

            Assert.Contains("部分上传将上传 0 个单元格，跳过 1 个单元格。", message);
            Assert.Contains("row-1 / status: 已跳过，单据已归档，禁止上传", message);
        }

        private static object CreateController(
            FakeSystemConnector connector,
            FakeWorksheetMetadataStore metadataStore,
            FakeDialogService dialogService,
            Func<string> sheetNameProvider)
        {
            return CreateController(connector, metadataStore, dialogService, sheetNameProvider, null);
        }

        private static object CreateController(
            FakeSystemConnector connector,
            FakeWorksheetMetadataStore metadataStore,
            FakeDialogService dialogService,
            Func<string> sheetNameProvider,
            Action authenticationLoginAction)
        {
            var addInAssembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var controllerType = addInAssembly.GetType("OfficeAgent.ExcelAddIn.RibbonSyncController", throwOnError: true);
            var executionService = CreateExecutionService(addInAssembly, connector, metadataStore);
            var dialogInterface = addInAssembly.GetType("OfficeAgent.ExcelAddIn.Dialogs.IRibbonSyncDialogService", throwOnError: true);
            var syncService = new WorksheetSyncService(
                new SystemConnectorRegistry(new ISystemConnector[] { connector }),
                metadataStore,
                new WorksheetChangeTracker(),
                new SyncOperationPreviewFactory());

            var ctorTypes = authenticationLoginAction == null
                ? new[]
                {
                    typeof(IWorksheetMetadataStore),
                    typeof(WorksheetSyncService),
                    typeof(Func<string>),
                    executionService.GetType(),
                    dialogInterface,
                }
                : new[]
                {
                    typeof(IWorksheetMetadataStore),
                    typeof(WorksheetSyncService),
                    typeof(Func<string>),
                    executionService.GetType(),
                    dialogInterface,
                    typeof(Action),
                };

            var ctor = controllerType.GetConstructor(
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                binder: null,
                types: ctorTypes,
                modifiers: null);

            if (ctor == null)
            {
                throw new InvalidOperationException("RibbonSyncController constructor with execution service was not found.");
            }

            return authenticationLoginAction == null
                ? ctor.Invoke(new object[] { metadataStore, syncService, sheetNameProvider, executionService, dialogService.GetTransparentProxy() })
                : ctor.Invoke(new object[] { metadataStore, syncService, sheetNameProvider, executionService, dialogService.GetTransparentProxy(), authenticationLoginAction });
        }

        private static object CreateExecutionService(
            Assembly addInAssembly,
            FakeSystemConnector connector,
            FakeWorksheetMetadataStore metadataStore)
        {
            var serviceType = addInAssembly.GetType("OfficeAgent.ExcelAddIn.WorksheetSyncExecutionService", throwOnError: true);
            var gridInterface = addInAssembly.GetType("OfficeAgent.ExcelAddIn.Excel.IWorksheetGridAdapter", throwOnError: true);
            var grid = new FakeWorksheetGridAdapter(gridInterface);
            var syncService = new WorksheetSyncService(
                new SystemConnectorRegistry(new ISystemConnector[] { connector }),
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

            return ctor.Invoke(new object[]
            {
                syncService,
                metadataStore,
                new FakeWorksheetSelectionReader(),
                grid.GetTransparentProxy(),
                new SyncOperationPreviewFactory(),
            });
        }

        private static void InvokeSelectProject(object controller, ProjectOption option)
        {
            var method = controller.GetType().GetMethod(
                "SelectProject",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                binder: null,
                types: new[] { typeof(ProjectOption) },
                modifiers: null);

            if (method == null)
            {
                throw new InvalidOperationException("RibbonSyncController.SelectProject(ProjectOption) was not found.");
            }

            method.Invoke(controller, new object[] { option });
        }

        private static void InvokeRefresh(object controller)
        {
            var method = controller.GetType().GetMethod(
                "RefreshActiveProjectFromSheetMetadata",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

            if (method == null)
            {
                throw new InvalidOperationException("RibbonSyncController.RefreshActiveProjectFromSheetMetadata() was not found.");
            }

            method.Invoke(controller, null);
        }

        private static void InvokeRefreshForSheet(object controller, string sheetName)
        {
            var method = controller.GetType().GetMethod(
                "RefreshProjectFromSheetMetadata",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                binder: null,
                types: new[] { typeof(string) },
                modifiers: null);

            if (method == null)
            {
                throw new InvalidOperationException("RibbonSyncController.RefreshProjectFromSheetMetadata(string) was not found.");
            }

            method.Invoke(controller, new object[] { sheetName });
        }

        private static void InvokeInvalidateRefreshState(object controller)
        {
            var method = controller.GetType().GetMethod(
                "InvalidateRefreshState",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

            if (method == null)
            {
                throw new InvalidOperationException("RibbonSyncController.InvalidateRefreshState() was not found.");
            }

            method.Invoke(controller, null);
        }

        private static void InvokeExecuteInitializeCurrentSheet(object controller)
        {
            var method = controller.GetType().GetMethod(
                "ExecuteInitializeCurrentSheet",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

            if (method == null)
            {
                throw new InvalidOperationException("RibbonSyncController.ExecuteInitializeCurrentSheet() was not found.");
            }

            method.Invoke(controller, null);
        }

        private static string ReadActiveProjectDisplayName(object controller)
        {
            return (string)controller.GetType().GetProperty(
                "ActiveProjectDisplayName",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic).GetValue(controller);
        }

        private static string ReadActiveProjectId(object controller)
        {
            return (string)controller.GetType().GetProperty(
                "ActiveProjectId",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic).GetValue(controller);
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

        private static string InvokeFormatProjectDropDownLabel(string projectId, string displayName)
        {
            var addInAssembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var ribbonType = addInAssembly.GetType("OfficeAgent.ExcelAddIn.AgentRibbon", throwOnError: true);
            var formatMethod = ribbonType.GetMethod(
                "FormatProjectDropDownLabel",
                BindingFlags.Static | BindingFlags.NonPublic);
            if (formatMethod == null)
            {
                throw new InvalidOperationException("AgentRibbon.FormatProjectDropDownLabel(string, string) was not found.");
            }

            return (string)formatMethod.Invoke(null, new object[] { projectId, displayName });
        }

        private static string InvokeBuildUploadPreviewInfoMessage(string operationName, SyncOperationPreview preview)
        {
            var addInAssembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var controllerType = addInAssembly.GetType("OfficeAgent.ExcelAddIn.RibbonSyncController", throwOnError: true);
            var method = controllerType.GetMethod(
                "BuildUploadPreviewInfoMessage",
                BindingFlags.Static | BindingFlags.NonPublic);
            if (method == null)
            {
                throw new InvalidOperationException("RibbonSyncController.BuildUploadPreviewInfoMessage was not found.");
            }

            return (string)method.Invoke(null, new object[] { operationName, preview });
        }

        private sealed class FakeSystemConnector : ISystemConnector
        {
            public FakeSystemConnector()
            {
                BindingSeed = new SheetBinding
                {
                    SheetName = "Sheet1",
                    SystemKey = "current-business-system",
                    ProjectId = "performance",
                    ProjectName = "绩效项目",
                    HeaderStartRow = 1,
                    HeaderRowCount = 2,
                    DataStartRow = 3,
                };
                FieldMappingDefinition = new FieldMappingTableDefinition
                {
                    SystemKey = "current-business-system",
                    Columns = new[]
                    {
                        new FieldMappingColumnDefinition { ColumnName = "ApiFieldKey", Role = FieldMappingSemanticRole.ApiFieldKey },
                    },
                };
                FieldMappingSeedRows = new[]
                {
                    new SheetFieldMappingRow
                    {
                        SheetName = "Sheet1",
                        Values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                        {
                            ["ApiFieldKey"] = "row_id",
                        },
                    },
                };
            }

            public string SystemKey => "current-business-system";

            public SheetBinding BindingSeed { get; set; }

            public FieldMappingTableDefinition FieldMappingDefinition { get; set; }

            public IReadOnlyList<SheetFieldMappingRow> FieldMappingSeedRows { get; set; }

            public string LastBuildFieldMappingSeedProjectId { get; private set; }

            public Exception BuildFieldMappingSeedException { get; set; }

            public IReadOnlyList<ProjectOption> GetProjects()
            {
                return Array.Empty<ProjectOption>();
            }

            public SheetBinding CreateBindingSeed(string sheetName, ProjectOption project)
            {
                return new SheetBinding
                {
                    SheetName = sheetName,
                    SystemKey = project?.SystemKey ?? string.Empty,
                    ProjectId = project?.ProjectId ?? string.Empty,
                    ProjectName = project?.DisplayName ?? string.Empty,
                    HeaderStartRow = BindingSeed.HeaderStartRow,
                    HeaderRowCount = BindingSeed.HeaderRowCount,
                    DataStartRow = BindingSeed.DataStartRow,
                };
            }

            public FieldMappingTableDefinition GetFieldMappingDefinition(string projectId)
            {
                return FieldMappingDefinition;
            }

            public IReadOnlyList<SheetFieldMappingRow> BuildFieldMappingSeed(string sheetName, string projectId)
            {
                if (BuildFieldMappingSeedException != null)
                {
                    throw BuildFieldMappingSeedException;
                }

                LastBuildFieldMappingSeedProjectId = projectId;
                return FieldMappingSeedRows;
            }

            public WorksheetSchema GetSchema(string projectId)
            {
                throw new NotSupportedException();
            }

            public IReadOnlyList<IDictionary<string, object>> Find(string projectId, IReadOnlyList<string> rowIds, IReadOnlyList<string> fieldKeys)
            {
                throw new NotSupportedException();
            }

            public void BatchSave(string projectId, IReadOnlyList<CellChange> changes)
            {
                throw new NotSupportedException();
            }
        }

        private sealed class FakeWorksheetMetadataStore : IWorksheetMetadataStore
        {
            public Dictionary<string, SheetBinding> Bindings { get; } = new Dictionary<string, SheetBinding>(StringComparer.OrdinalIgnoreCase);

            public Dictionary<string, SheetFieldMappingRow[]> FieldMappings { get; } = new Dictionary<string, SheetFieldMappingRow[]>(StringComparer.OrdinalIgnoreCase);

            public int LoadBindingCallCount { get; private set; }

            public SheetBinding LastSavedBinding { get; private set; }

            public SheetFieldMappingRow[] LastSavedFieldMappings { get; private set; } = Array.Empty<SheetFieldMappingRow>();

            public void SaveBinding(SheetBinding binding)
            {
                LastSavedBinding = binding;
                Bindings[binding.SheetName] = binding;
            }

            public SheetBinding LoadBinding(string sheetName)
            {
                LoadBindingCallCount++;
                if (!Bindings.TryGetValue(sheetName, out var binding))
                {
                    throw new InvalidOperationException("No binding.");
                }

                return binding;
            }

            public void SaveFieldMappings(string sheetName, FieldMappingTableDefinition definition, IReadOnlyList<SheetFieldMappingRow> rows)
            {
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

        private sealed class FakeDialogService : RealProxy
        {
            public FakeDialogService()
                : base(LoadDialogInterfaceType())
            {
            }

            public List<string> InfoMessages { get; } = new List<string>();

            public List<string> WarningMessages { get; } = new List<string>();

            public List<string> ErrorMessages { get; } = new List<string>();

            public List<SheetBinding> ProjectLayoutPrompts { get; } = new List<SheetBinding>();

            public List<string> AuthenticationRequiredMessages { get; } = new List<string>();

            public SheetBinding NextProjectLayoutBinding { get; set; }

            public bool AuthenticationRequiredResult { get; set; }

            public override IMessage Invoke(IMessage msg)
            {
                var call = (IMethodCallMessage)msg;
                switch (call.MethodName)
                {
                    case "ConfirmDownload":
                    case "ConfirmUpload":
                        return new ReturnMessage(true, null, 0, call.LogicalCallContext, call);
                    case "ShowProjectLayoutDialog":
                        ProjectLayoutPrompts.Add(CloneBinding((SheetBinding)call.InArgs[0]));
                        return new ReturnMessage(CloneBinding(NextProjectLayoutBinding), null, 0, call.LogicalCallContext, call);
                    case "ShowInfo":
                        InfoMessages.Add((string)call.InArgs[0]);
                        return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
                    case "ShowWarning":
                        WarningMessages.Add((string)call.InArgs[0]);
                        return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
                    case "ShowError":
                        ErrorMessages.Add((string)call.InArgs[0]);
                        return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
                    case "ShowAuthenticationRequired":
                        AuthenticationRequiredMessages.Add((string)call.InArgs[0]);
                        return new ReturnMessage(AuthenticationRequiredResult, null, 0, call.LogicalCallContext, call);
                    default:
                        throw new NotSupportedException(call.MethodName);
                }
            }

            public new object GetTransparentProxy()
            {
                return base.GetTransparentProxy();
            }

            private static Type LoadDialogInterfaceType()
            {
                return Assembly.LoadFrom(ResolveAddInAssemblyPath())
                    .GetType("OfficeAgent.ExcelAddIn.Dialogs.IRibbonSyncDialogService", throwOnError: true);
            }

            private static SheetBinding CloneBinding(SheetBinding binding)
            {
                if (binding == null)
                {
                    return null;
                }

                return new SheetBinding
                {
                    SheetName = binding.SheetName,
                    SystemKey = binding.SystemKey,
                    ProjectId = binding.ProjectId,
                    ProjectName = binding.ProjectName,
                    HeaderStartRow = binding.HeaderStartRow,
                    HeaderRowCount = binding.HeaderRowCount,
                    DataStartRow = binding.DataStartRow,
                };
            }
        }

        private sealed class FakeWorksheetSelectionReader : IWorksheetSelectionReader
        {
            public IReadOnlyList<SelectedVisibleCell> ReadVisibleSelection()
            {
                return Array.Empty<SelectedVisibleCell>();
            }
        }

        private sealed class FakeWorksheetGridAdapter : RealProxy
        {
            public FakeWorksheetGridAdapter(Type interfaceType)
                : base(interfaceType)
            {
            }

            public override IMessage Invoke(IMessage msg)
            {
                var call = (IMethodCallMessage)msg;
                switch (call.MethodName)
                {
                    case "GetCellText":
                        return new ReturnMessage(string.Empty, null, 0, call.LogicalCallContext, call);
                    case "GetLastUsedRow":
                    case "GetLastUsedColumn":
                        return new ReturnMessage(0, null, 0, call.LogicalCallContext, call);
                    case "SetCellText":
                    case "ClearRange":
                    case "ClearWorksheet":
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
        }
    }
}
