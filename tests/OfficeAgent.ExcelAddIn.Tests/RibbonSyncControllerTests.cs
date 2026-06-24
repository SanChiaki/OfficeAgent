using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using System.Threading;
using System.Threading.Tasks;
using OfficeAgent.Core;
using OfficeAgent.Core.Analytics;
using OfficeAgent.Core.Diagnostics;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
using OfficeAgent.Core.Sync;
using OfficeAgent.ExcelAddIn.Dialogs;
using OfficeAgent.ExcelAddIn.Excel;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    [CollectionDefinition(Name)]
    public sealed class OfficeAgentLogCollection
    {
        public const string Name = "OfficeAgentLog";
    }

    [Collection(OfficeAgentLogCollection.Name)]
    public sealed class RibbonSyncControllerTests
    {
        [Fact]
        public void NewControllerDefaultsToSelectProjectDisplayWhenNoBinding()
        {
            var controller = CreateController(new FakeSystemConnector(), new FakeWorksheetMetadataStore(), new FakeDialogService(), () => "Sheet1");

            Assert.Equal("Select project", ReadActiveProjectDisplayName(controller));
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
        public void SelectProjectCancelWritesProjectLayoutDiagnostics()
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

            var logs = CaptureLogEntries(() => InvokeSelectProject(controller, new ProjectOption
            {
                SystemKey = "current-business-system",
                ProjectId = "new-project",
                DisplayName = "新项目",
            }));

            Assert.Contains(logs, entry =>
                entry.Level == "info" &&
                entry.Component == "ribbon_sync" &&
                entry.EventName == "project.select.begin" &&
                entry.Details.Contains("SheetName=Sheet1") &&
                entry.Details.Contains("TargetProjectId=new-project") &&
                entry.Details.Contains("ExistingProjectId=old-project"));
            Assert.Contains(logs, entry =>
                entry.Level == "info" &&
                entry.Component == "ribbon_sync" &&
                entry.EventName == "project.layout_dialog.show" &&
                entry.Details.Contains("HeaderStartRow=5") &&
                entry.Details.Contains("DataStartRow=7"));
            Assert.Contains(logs, entry =>
                entry.Level == "warn" &&
                entry.Component == "ribbon_sync" &&
                entry.EventName == "project.layout_dialog.cancelled" &&
                entry.Details.Contains("RestoredProjectId=old-project"));
        }

        [Fact]
        public void SelectProjectConfirmedWritesBindingSaveDiagnostics()
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

            var logs = CaptureLogEntries(() => InvokeSelectProject(controller, new ProjectOption
            {
                SystemKey = "current-business-system",
                ProjectId = "performance",
                DisplayName = "绩效项目",
            }));

            Assert.Contains(logs, entry =>
                entry.Level == "info" &&
                entry.Component == "ribbon_sync" &&
                entry.EventName == "project.binding.save.begin" &&
                entry.Details.Contains("SheetName=Sheet1") &&
                entry.Details.Contains("TargetProjectId=performance") &&
                entry.Details.Contains("HeaderStartRow=4"));
            Assert.Contains(logs, entry =>
                entry.Level == "info" &&
                entry.Component == "ribbon_sync" &&
                entry.EventName == "project.binding.save.completed" &&
                entry.Details.Contains("SheetName=Sheet1") &&
                entry.Details.Contains("TargetProjectId=performance") &&
                entry.Details.Contains("DataStartRow=5"));
        }

        [Fact]
        public void SelectProjectSaveBindingFailureWritesErrorDiagnostics()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore
            {
                SaveBindingException = new InvalidOperationException("metadata write failed"),
            };
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

            var capture = CaptureLogEntriesAllowingFailure(() => InvokeSelectProject(controller, new ProjectOption
            {
                SystemKey = "current-business-system",
                ProjectId = "performance",
                DisplayName = "绩效项目",
            }));

            var failure = Assert.IsType<TargetInvocationException>(capture.Failure);
            Assert.IsType<InvalidOperationException>(failure.InnerException);
            Assert.Contains(capture.Entries, entry =>
                entry.Level == "error" &&
                entry.Component == "ribbon_sync" &&
                entry.EventName == "project.binding.save.failed" &&
                entry.Details.Contains("SheetName=Sheet1") &&
                entry.Details.Contains("TargetProjectId=performance") &&
                entry.Exception.Contains("metadata write failed"));
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
            Assert.Contains(dialogService.InfoMessages, message => message.IndexOf("current sheet content was not changed", StringComparison.OrdinalIgnoreCase) >= 0);
        }

        [Fact]
        public void ExecuteInitializeCurrentSheetTracksCompletedEventWithProjectName()
        {
            var connector = new FakeSystemConnector();
            var analytics = new RecordingAnalyticsService();
            var controller = CreateController(
                connector,
                activeSheetName: "Sheet1",
                analyticsService: analytics);

            InvokeSelectProject(controller, new ProjectOption
            {
                SystemKey = connector.SystemKey,
                ProjectId = "performance",
                DisplayName = "绩效项目",
            });

            InvokeExecuteInitializeCurrentSheet(controller);

            Assert.Contains(analytics.Events, analyticsEvent =>
                analyticsEvent.EventName == "ribbon.initialize.completed" &&
                Equals(analyticsEvent.Properties["projectId"], "performance") &&
                Equals(analyticsEvent.Properties["projectName"], "绩效项目") &&
                Equals(analyticsEvent.Properties["sheetName"], "Sheet1"));
        }

        [Fact]
        public void ExecuteInitializeCurrentSheetConfigOnlyShowsDialogAndDoesNotModifyBusinessCells()
        {
            var connector = new FakeBusinessTemplateConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var dialogService = new FakeDialogService
            {
                NextInitializeSheetResult = new InitializeSheetDialogResult
                {
                    Mode = InitializeSheetMode.ConfigOnly,
                },
            };
            metadataStore.Bindings["Sheet1"] = new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = connector.SystemKey,
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 5,
                HeaderRowCount = 2,
                DataStartRow = 9,
            };

            var (controller, grid) = CreateControllerWithGrid(
                connector,
                metadataStore,
                dialogService,
                () => "Sheet1",
                aiClient: null);
            grid.SetCell("Sheet1", 1, 1, "keep me");
            InvokeRefresh(controller);

            InvokeExecuteInitializeCurrentSheet(controller);

            var request = Assert.Single(dialogService.InitializeSheetRequests);
            Assert.Equal("绩效项目", request.ProjectDisplayName);
            Assert.True(request.IsBlankSheet);
            Assert.True(request.SupportsTemplateImport);
            Assert.Single(dialogService.InitializeSheetTemplateLoadResults);
            Assert.Equal("standard", dialogService.InitializeSheetTemplateLoadResults[0].Templates[0].TemplateId);
            Assert.Equal("keep me", grid.GetCell("Sheet1", 1, 1));
            Assert.Equal(0, grid.ClearRangeCallCount);
            Assert.Equal(0, grid.WriteRangeValuesCallCount);
            Assert.Equal("performance", connector.LastBuildFieldMappingSeedProjectId);
            Assert.Equal(5, metadataStore.LastSavedBinding.HeaderStartRow);
            Assert.Contains(dialogService.InfoMessages, message => message.IndexOf("current sheet content was not changed", StringComparison.OrdinalIgnoreCase) >= 0);
            Assert.DoesNotContain(dialogService.InfoMessages, message => message.IndexOf("template", StringComparison.OrdinalIgnoreCase) >= 0);
        }

        [Fact]
        public void ExecuteInitializeCurrentSheetTemplateImportRunsProgressAndShowsTemplateSuccess()
        {
            var connector = new FakeBusinessTemplateConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var dialogService = new FakeDialogService
            {
                NextInitializeSheetResult = new InitializeSheetDialogResult
                {
                    Mode = InitializeSheetMode.TemplateImport,
                    SelectedTemplate = new BusinessExportTemplateOption
                    {
                        TemplateId = "standard",
                        TemplateName = "标准作业表",
                    },
                },
            };
            metadataStore.Bindings["Sheet1"] = CreateInitializedBinding(connector.SystemKey);

            var (controller, _) = CreateControllerWithGrid(
                connector,
                metadataStore,
                dialogService,
                () => "Sheet1",
                aiClient: null);
            InvokeRefresh(controller);

            InvokeExecuteInitializeCurrentSheet(controller);

            Assert.Equal(1, dialogService.InitializeTemplateImportProgressRunCount);
            Assert.Contains("downloading", dialogService.InitializeTemplateImportProgressCalls);
            Assert.Contains("importing", dialogService.InitializeTemplateImportProgressCalls);
            Assert.Contains("writingConfiguration", dialogService.InitializeTemplateImportProgressCalls);
            Assert.Equal("performance", connector.LastExportProjectId);
            Assert.Equal("standard", connector.LastExportTemplateId);
            Assert.Contains(dialogService.InfoMessages, message => message.IndexOf("created from the template", StringComparison.OrdinalIgnoreCase) >= 0);
        }

        [Theory]
        [InlineData("xISDP_Setting")]
        [InlineData("xISDP_Log")]
        [InlineData("xisdp_log")]
        public void ExecuteInitializeCurrentSheetBlocksManagedSheetsBeforeDialog(string sheetName)
        {
            var connector = new FakeBusinessTemplateConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var dialogService = new FakeDialogService
            {
                NextInitializeSheetResult = new InitializeSheetDialogResult
                {
                    Mode = InitializeSheetMode.ConfigOnly,
                },
            };
            metadataStore.Bindings[sheetName] = new SheetBinding
            {
                SheetName = sheetName,
                SystemKey = connector.SystemKey,
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 1,
                HeaderRowCount = 2,
                DataStartRow = 3,
            };
            var controller = CreateController(connector, metadataStore, dialogService, () => sheetName);
            InvokeRefresh(controller);

            InvokeExecuteInitializeCurrentSheet(controller);

            Assert.Empty(dialogService.InitializeSheetRequests);
            Assert.Empty(dialogService.InitializeSheetTemplateLoadResults);
            Assert.Empty(dialogService.InfoMessages);
            var warning = Assert.Single(dialogService.WarningMessages);
            Assert.Contains("xISDP_Setting", warning, StringComparison.Ordinal);
            Assert.Contains("xISDP_Log", warning, StringComparison.Ordinal);
            Assert.Null(metadataStore.LastSavedBinding);
            Assert.Null(connector.LastBuildFieldMappingSeedProjectId);
        }

        [Theory]
        [InlineData("xISDP_Setting")]
        [InlineData("xISDP_Log")]
        public void ExecuteInitializeCurrentSheetBlocksManagedSheetsBeforeProjectSelectionWarning(string sheetName)
        {
            var connector = new FakeBusinessTemplateConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var dialogService = new FakeDialogService();
            var controller = CreateController(connector, metadataStore, dialogService, () => sheetName);

            InvokeExecuteInitializeCurrentSheet(controller);

            Assert.Empty(dialogService.InitializeSheetRequests);
            Assert.Empty(dialogService.InitializeSheetTemplateLoadResults);
            Assert.Empty(dialogService.InfoMessages);
            var warning = Assert.Single(dialogService.WarningMessages);
            Assert.Contains("xISDP_Setting", warning, StringComparison.Ordinal);
            Assert.Contains("xISDP_Log", warning, StringComparison.Ordinal);
            Assert.DoesNotContain("Select a project first.", dialogService.WarningMessages);
            Assert.Null(metadataStore.LastSavedBinding);
            Assert.Null(connector.LastBuildFieldMappingSeedProjectId);
        }

        [Fact]
        public void ExecuteInitializeCurrentSheetTemplateProgressCancellationIsQuiet()
        {
            var connector = new FakeBusinessTemplateConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var dialogService = new FakeDialogService
            {
                CancelInitializeTemplateImportProgress = true,
                NextInitializeSheetResult = new InitializeSheetDialogResult
                {
                    Mode = InitializeSheetMode.TemplateImport,
                    SelectedTemplate = new BusinessExportTemplateOption
                    {
                        TemplateId = "standard",
                        TemplateName = "标准作业表",
                    },
                },
            };
            metadataStore.Bindings["Sheet1"] = CreateInitializedBinding(connector.SystemKey);

            var (controller, _) = CreateControllerWithGrid(
                connector,
                metadataStore,
                dialogService,
                () => "Sheet1",
                aiClient: null);
            InvokeRefresh(controller);

            InvokeExecuteInitializeCurrentSheet(controller);

            Assert.Equal(1, dialogService.InitializeTemplateImportProgressRunCount);
            Assert.Empty(dialogService.ErrorMessages);
            Assert.Empty(dialogService.InfoMessages);
        }

        [Fact]
        public void ExecuteInitializeCurrentSheetCanceledDialogTracksCanceledAndShowsNoMessages()
        {
            var connector = new FakeBusinessTemplateConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var dialogService = new FakeDialogService();
            var analytics = new RecordingAnalyticsService();
            dialogService.CancelInitializeSheetDialog = true;
            metadataStore.Bindings["Sheet1"] = CreateInitializedBinding(connector.SystemKey);
            var controller = CreateController(
                connector,
                metadataStore,
                dialogService,
                () => "Sheet1",
                authenticationLoginAction: null,
                analyticsService: analytics);
            InvokeRefresh(controller);

            InvokeExecuteInitializeCurrentSheet(controller);

            Assert.Single(dialogService.InitializeSheetRequests);
            Assert.Empty(dialogService.InfoMessages);
            Assert.Empty(dialogService.ErrorMessages);
            Assert.Contains(analytics.Events, analyticsEvent =>
                analyticsEvent.EventName == "ribbon.initialize.canceled" &&
                Equals(analyticsEvent.Properties["projectId"], "performance") &&
                Equals(analyticsEvent.Properties["sheetName"], "Sheet1"));
        }

        [Fact]
        public void ExecuteInitializeCurrentSheetTracksTemplateImportAnalytics()
        {
            var connector = new FakeBusinessTemplateConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var dialogService = new FakeDialogService
            {
                NextInitializeSheetResult = new InitializeSheetDialogResult
                {
                    Mode = InitializeSheetMode.TemplateImport,
                    SelectedTemplate = new BusinessExportTemplateOption
                    {
                        TemplateId = "standard",
                        TemplateName = "标准作业表",
                    },
                },
            };
            var analytics = new RecordingAnalyticsService();
            metadataStore.Bindings["Sheet1"] = CreateInitializedBinding(connector.SystemKey);
            var (controller, _) = CreateControllerWithGrid(
                connector,
                metadataStore,
                dialogService,
                () => "Sheet1",
                aiClient: null,
                selectionReader: null,
                analyticsService: analytics);
            InvokeRefresh(controller);

            InvokeExecuteInitializeCurrentSheet(controller);

            var started = Assert.Single(analytics.Events, analyticsEvent => analyticsEvent.EventName == "ribbon.initialize_template_import.started");
            AssertTemplateImportProperties(started, templateId: "standard", templateName: "标准作业表", isBlankSheet: true);
            Assert.True(Convert.ToInt64(started.Properties["durationMs"]) >= 0);

            var completed = Assert.Single(analytics.Events, analyticsEvent => analyticsEvent.EventName == "ribbon.initialize_template_import.completed");
            AssertTemplateImportProperties(completed, templateId: "standard", templateName: "标准作业表", isBlankSheet: true);
            Assert.True(Convert.ToInt64(completed.Properties["durationMs"]) >= 0);
        }

        [Fact]
        public void ExecuteInitializeCurrentSheetTracksTemplateImportCanceledAnalytics()
        {
            var connector = new FakeBusinessTemplateConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var dialogService = new FakeDialogService
            {
                CancelInitializeTemplateImportProgress = true,
                NextInitializeSheetResult = new InitializeSheetDialogResult
                {
                    Mode = InitializeSheetMode.TemplateImport,
                    SelectedTemplate = new BusinessExportTemplateOption
                    {
                        TemplateId = "standard",
                        TemplateName = "标准作业表",
                    },
                },
            };
            var analytics = new RecordingAnalyticsService();
            metadataStore.Bindings["Sheet1"] = CreateInitializedBinding(connector.SystemKey);
            var (controller, _) = CreateControllerWithGrid(
                connector,
                metadataStore,
                dialogService,
                () => "Sheet1",
                aiClient: null,
                selectionReader: null,
                analyticsService: analytics);
            InvokeRefresh(controller);

            InvokeExecuteInitializeCurrentSheet(controller);

            var canceled = Assert.Single(analytics.Events, analyticsEvent => analyticsEvent.EventName == "ribbon.initialize_template_import.canceled");
            AssertTemplateImportProperties(canceled, templateId: "standard", templateName: "标准作业表", isBlankSheet: true);
            Assert.True(Convert.ToInt64(canceled.Properties["durationMs"]) >= 0);
            Assert.DoesNotContain(analytics.Events, analyticsEvent => analyticsEvent.EventName == "ribbon.initialize_template_import.completed");
        }

        [Fact]
        public void ExecuteInitializeCurrentSheetTracksTemplateImportFailedAnalyticsWithStageAndExceptionType()
        {
            var connector = new FakeBusinessTemplateConnector
            {
                ExportException = new InvalidOperationException("download failed"),
            };
            var metadataStore = new FakeWorksheetMetadataStore();
            var dialogService = new FakeDialogService
            {
                NextInitializeSheetResult = new InitializeSheetDialogResult
                {
                    Mode = InitializeSheetMode.TemplateImport,
                    SelectedTemplate = new BusinessExportTemplateOption
                    {
                        TemplateId = "standard",
                        TemplateName = "标准作业表",
                    },
                },
            };
            var analytics = new RecordingAnalyticsService();
            metadataStore.Bindings["Sheet1"] = CreateInitializedBinding(connector.SystemKey);
            var (controller, _) = CreateControllerWithGrid(
                connector,
                metadataStore,
                dialogService,
                () => "Sheet1",
                aiClient: null,
                selectionReader: null,
                analyticsService: analytics);
            InvokeRefresh(controller);

            InvokeExecuteInitializeCurrentSheet(controller);

            var failed = Assert.Single(analytics.Events, analyticsEvent => analyticsEvent.EventName == "ribbon.initialize_template_import.failed");
            AssertTemplateImportProperties(failed, templateId: "standard", templateName: "标准作业表", isBlankSheet: true);
            Assert.Equal("initialize_template_import", failed.Properties["failedStage"]);
            Assert.Equal("InvalidOperationException", failed.Properties["exceptionType"]);
            Assert.True(Convert.ToInt64(failed.Properties["durationMs"]) >= 0);
            Assert.NotNull(failed.Error);
            Assert.Equal("InvalidOperationException", failed.Error.ExceptionType);
            Assert.Single(dialogService.ErrorMessages);
        }

        [Fact]
        public void ExecuteInitializeCurrentSheetFailureWritesDiagnostics()
        {
            var connector = new FakeSystemConnector
            {
                BuildFieldMappingSeedException = new TaskCanceledException("A task was canceled."),
            };
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

            var logs = CaptureLogEntries(() => InvokeExecuteInitializeCurrentSheet(controller));

            Assert.Single(dialogService.ErrorMessages);
            Assert.Equal("A task was canceled.", dialogService.ErrorMessages[0]);
            Assert.Contains(logs, entry =>
                entry.Level == "info" &&
                entry.Component == "ribbon_sync" &&
                entry.EventName == "initialize_sheet.begin" &&
                entry.Details.Contains("SheetName=Sheet1") &&
                entry.Details.Contains("SystemKey=current-business-system") &&
                entry.Details.Contains("ProjectId=performance") &&
                entry.Details.Contains("ProjectName=绩效项目"));
            Assert.Contains(logs, entry =>
                entry.Level == "error" &&
                entry.Component == "ribbon_sync" &&
                entry.EventName == "initialize_sheet.failed" &&
                entry.Details.Contains("SheetName=Sheet1") &&
                entry.Details.Contains("ProjectId=performance") &&
                entry.Exception.Contains("TaskCanceledException") &&
                entry.Exception.Contains("A task was canceled."));
        }

        [Fact]
        public void ExecuteInitializeCurrentSheetLocalizesLoginPromptWhenAuthenticationIsRequired()
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
            Assert.Equal("You're not signed in. Sign in first.", dialogService.AuthenticationRequiredMessages[0]);
            Assert.True(loginTriggered);
            Assert.Empty(dialogService.ErrorMessages);
        }

        [Fact]
        public void ExecuteInitializeCurrentSheetNotifiesAuthenticationRequiredBeforePrompt()
        {
            var connector = new FakeSystemConnector
            {
                BuildFieldMappingSeedException = new AuthenticationRequiredException("当前未登录，请先登录"),
            };
            var metadataStore = new FakeWorksheetMetadataStore();
            var dialogService = new FakeDialogService
            {
                AuthenticationRequiredResult = false,
            };
            var authenticationRequiredCount = 0;
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
                authenticationLoginAction: null,
                authenticationRequiredAction: () => authenticationRequiredCount++);
            InvokeRefresh(controller);

            InvokeExecuteInitializeCurrentSheet(controller);

            Assert.Equal(1, authenticationRequiredCount);
            Assert.Single(dialogService.AuthenticationRequiredMessages);
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

            Assert.Equal("Select project", ReadActiveProjectDisplayName(controller));
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

            var controller = CreateController(new FakeSystemConnector(), metadataStore, new FakeDialogService(), () => "xISDP_Setting");

            InvokeRefresh(controller);

            Assert.Equal(string.Empty, ReadActiveProjectId(controller));
            Assert.Equal("Select project", ReadActiveProjectDisplayName(controller));
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
            InvokeRefreshForSheet(controller, "xISDP_Setting");

            Assert.Equal(string.Empty, ReadActiveProjectId(controller));
            Assert.Equal("Select project", ReadActiveProjectDisplayName(controller));
        }

        [Fact]
        public void RefreshActiveProjectFromSettingsSheetLoadsBindingWhenSettingsSheetIsExplicitlyBound()
        {
            var metadataStore = new FakeWorksheetMetadataStore();
            metadataStore.Bindings["xISDP_Setting"] = new SheetBinding
            {
                SheetName = "xISDP_Setting",
                SystemKey = "current-business-system",
                ProjectId = "settings-project",
                ProjectName = "设置页项目",
            };

            var controller = CreateController(new FakeSystemConnector(), metadataStore, new FakeDialogService(), () => "xISDP_Setting");

            InvokeRefresh(controller);

            Assert.Equal("settings-project", ReadActiveProjectId(controller));
            Assert.Equal("设置页项目", ReadActiveProjectDisplayName(controller));
        }

        [Fact]
        public void SelectProjectFromSettingsSheetRefreshesMetadataPresentationAfterSavingBinding()
        {
            var connector = new FakeSystemConnector();
            var metadataStore = new FakeWorksheetMetadataStore();
            var dialogService = new FakeDialogService
            {
                NextProjectLayoutBinding = new SheetBinding
                {
                    SheetName = "xISDP_Setting",
                    SystemKey = "current-business-system",
                    ProjectId = "settings-project",
                    ProjectName = "设置页项目",
                    HeaderStartRow = 4,
                    HeaderRowCount = 1,
                    DataStartRow = 5,
                },
            };
            var controller = CreateController(connector, metadataStore, dialogService, () => "xISDP_Setting");

            InvokeSelectProject(controller, new ProjectOption
            {
                SystemKey = "current-business-system",
                ProjectId = "settings-project",
                DisplayName = "设置页项目",
            });

            Assert.Equal("xISDP_Setting", metadataStore.LastRefreshedPresentationSheetName);
            Assert.False(metadataStore.LastRefreshedPresentationHideTemplateBindingRows);
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
        public void RibbonSyncControllerRoutesDownloadAndUploadStatusMessagesThroughHostLocalizedStrings()
        {
            var controllerText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "RibbonSyncController.cs"));

            Assert.Contains("LocalizeSyncOperationName(", controllerText, StringComparison.Ordinal);
            Assert.Contains("FormatDownloadCompletedMessage(", controllerText, StringComparison.Ordinal);
            Assert.Contains("FormatUploadNoChangesMessage(", controllerText, StringComparison.Ordinal);
            Assert.Contains("FormatUploadCompletedMessage(", controllerText, StringComparison.Ordinal);
            Assert.DoesNotContain("没有可提交的单元格。", controllerText, StringComparison.Ordinal);
        }

        [Fact]
        public void BuildUploadPreviewInfoMessageIncludesSkippedReasonsWhenNothingWillUpload()
        {
            var preview = new SyncOperationPreview
            {
                Summary = "Upload will submit 0 cell(s) and skip 1 cell(s).",
                Details = new[] { "row-1 / status: Skipped, 单据已归档，禁止上传" },
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

            var message = InvokeBuildUploadPreviewInfoMessage("Upload", preview);

            Assert.Contains("Upload will submit 0 cell(s) and skip 1 cell(s).", message);
            Assert.Contains("row-1 / status: Skipped, 单据已归档，禁止上传", message);
        }

        [Fact]
        public void ExecuteAiColumnMappingConfirmsPreviewBeforeSavingMappings()
        {
            var connector = new FakeSystemConnector
            {
                FieldMappingDefinition = BuildAiMappingDefinition(),
            };
            var metadataStore = new FakeWorksheetMetadataStore();
            var dialogService = new FakeDialogService { AiColumnMappingConfirmResult = true };
            var aiClient = new FakeAiColumnMappingClient
            {
                Response = new AiColumnMappingResponse
                {
                    Mappings = new[]
                    {
                        new AiColumnMappingSuggestion
                        {
                            ExcelColumn = 2,
                            ActualL1 = "项目负责人",
                            TargetHeaderId = "owner_name",
                            TargetApiFieldKey = "owner_name",
                            Confidence = 0.91,
                        },
                    },
                },
            };
            metadataStore.Bindings["Sheet1"] = CreateAiMappingBinding();
            metadataStore.FieldMappings["Sheet1"] = BuildAiMappings("Sheet1");
            var (controller, grid) = CreateControllerWithGrid(
                connector,
                metadataStore,
                dialogService,
                () => "Sheet1",
                aiClient);
            grid.SetCell("Sheet1", 3, 1, "ID");
            grid.SetCell("Sheet1", 3, 2, "项目负责人");

            InvokeRefresh(controller);
            InvokeExecuteAiColumnMapping(controller);

            Assert.Single(dialogService.AiColumnMappingPreviews);
            Assert.Equal("Sheet1", aiClient.LastRequest.SheetName);
            Assert.Equal("项目负责人", aiClient.LastRequest.ActualHeaders.Single(header => header.ExcelColumn == 2).ActualL1);
            Assert.Equal("项目负责人", metadataStore.LastSavedFieldMappings.Single(row => row.Values["HeaderId"] == "owner_name").Values["CurrentL1"]);
            Assert.Contains(dialogService.InfoMessages, message => message.IndexOf("Applied: 1", StringComparison.Ordinal) >= 0);
            Assert.Empty(dialogService.ErrorMessages);
        }

        [Fact]
        public void ExecuteAiColumnMappingDoesNotSaveWhenPreviewIsCancelled()
        {
            var connector = new FakeSystemConnector
            {
                FieldMappingDefinition = BuildAiMappingDefinition(),
            };
            var metadataStore = new FakeWorksheetMetadataStore();
            var dialogService = new FakeDialogService { AiColumnMappingConfirmResult = false };
            var aiClient = new FakeAiColumnMappingClient
            {
                Response = new AiColumnMappingResponse
                {
                    Mappings = new[]
                    {
                        new AiColumnMappingSuggestion
                        {
                            ExcelColumn = 2,
                            ActualL1 = "项目负责人",
                            TargetHeaderId = "owner_name",
                            TargetApiFieldKey = "owner_name",
                            Confidence = 0.91,
                        },
                    },
                },
            };
            metadataStore.Bindings["Sheet1"] = CreateAiMappingBinding();
            metadataStore.FieldMappings["Sheet1"] = BuildAiMappings("Sheet1");
            var (controller, grid) = CreateControllerWithGrid(
                connector,
                metadataStore,
                dialogService,
                () => "Sheet1",
                aiClient);
            grid.SetCell("Sheet1", 3, 1, "ID");
            grid.SetCell("Sheet1", 3, 2, "项目负责人");

            InvokeRefresh(controller);
            InvokeExecuteAiColumnMapping(controller);

            Assert.Single(dialogService.AiColumnMappingPreviews);
            Assert.Empty(metadataStore.LastSavedFieldMappings);
            Assert.Empty(dialogService.InfoMessages);
        }

        [Fact]
        public void ExecuteAiColumnMappingDoesNotShowPreviewWhenNoAcceptedMappingsExist()
        {
            var connector = new FakeSystemConnector
            {
                FieldMappingDefinition = BuildAiMappingDefinition(),
            };
            var metadataStore = new FakeWorksheetMetadataStore();
            var dialogService = new FakeDialogService { AiColumnMappingConfirmResult = true };
            var aiClient = new FakeAiColumnMappingClient
            {
                Response = new AiColumnMappingResponse
                {
                    Unmatched = new[]
                    {
                        new AiColumnMappingUnmatchedHeader
                        {
                            ExcelColumn = 2,
                            ActualL1 = "项目负责人",
                            Reason = "No clear match",
                        },
                    },
                },
            };
            metadataStore.Bindings["Sheet1"] = CreateAiMappingBinding();
            metadataStore.FieldMappings["Sheet1"] = BuildAiMappings("Sheet1");
            var (controller, grid) = CreateControllerWithGrid(
                connector,
                metadataStore,
                dialogService,
                () => "Sheet1",
                aiClient);
            grid.SetCell("Sheet1", 3, 1, "ID");
            grid.SetCell("Sheet1", 3, 2, "项目负责人");

            InvokeRefresh(controller);
            InvokeExecuteAiColumnMapping(controller);

            Assert.Empty(dialogService.AiColumnMappingPreviews);
            Assert.Empty(metadataStore.LastSavedFieldMappings);
            Assert.Contains(dialogService.InfoMessages, message => message.IndexOf("no accepted mappings", StringComparison.OrdinalIgnoreCase) >= 0);
        }

        [Fact]
        public void ExecuteAiColumnMappingSkipsMappingsUncheckedInPreview()
        {
            var connector = new FakeSystemConnector
            {
                FieldMappingDefinition = BuildAiMappingDefinition(),
            };
            var metadataStore = new FakeWorksheetMetadataStore();
            var dialogService = new FakeDialogService
            {
                AiColumnMappingConfirmResult = true,
                OnConfirmAiColumnMapping = preview =>
                {
                    preview.Items.Single().ShouldApply = false;
                },
            };
            var aiClient = new FakeAiColumnMappingClient
            {
                Response = new AiColumnMappingResponse
                {
                    Mappings = new[]
                    {
                        new AiColumnMappingSuggestion
                        {
                            ExcelColumn = 2,
                            ActualL1 = "项目负责人",
                            TargetHeaderId = "owner_name",
                            TargetApiFieldKey = "owner_name",
                            Confidence = 0.91,
                        },
                    },
                },
            };
            metadataStore.Bindings["Sheet1"] = CreateAiMappingBinding();
            metadataStore.FieldMappings["Sheet1"] = BuildAiMappings("Sheet1");
            var (controller, grid) = CreateControllerWithGrid(
                connector,
                metadataStore,
                dialogService,
                () => "Sheet1",
                aiClient);
            grid.SetCell("Sheet1", 3, 1, "ID");
            grid.SetCell("Sheet1", 3, 2, "项目负责人");

            InvokeRefresh(controller);
            InvokeExecuteAiColumnMapping(controller);

            Assert.Single(dialogService.AiColumnMappingPreviews);
            Assert.Equal("负责人", metadataStore.LastSavedFieldMappings.Single(row => row.Values["HeaderId"] == "owner_name").Values["CurrentL1"]);
            Assert.Contains(dialogService.InfoMessages, message => message.IndexOf("no accepted mappings", StringComparison.OrdinalIgnoreCase) >= 0);
        }

        [Fact]
        public void ExecuteAiColumnMappingRunsModelCallThroughCancellableProgressDialog()
        {
            var connector = new FakeSystemConnector
            {
                FieldMappingDefinition = BuildAiMappingDefinition(),
            };
            var metadataStore = new FakeWorksheetMetadataStore();
            var dialogService = new FakeDialogService { AiColumnMappingConfirmResult = true };
            var aiClient = new FakeAiColumnMappingClient
            {
                Response = new AiColumnMappingResponse
                {
                    Mappings = new[]
                    {
                        new AiColumnMappingSuggestion
                        {
                            ExcelColumn = 2,
                            ActualL1 = "项目负责人",
                            TargetHeaderId = "owner_name",
                            TargetApiFieldKey = "owner_name",
                            Confidence = 0.91,
                        },
                    },
                },
            };
            metadataStore.Bindings["Sheet1"] = CreateAiMappingBinding();
            metadataStore.FieldMappings["Sheet1"] = BuildAiMappings("Sheet1");
            var (controller, grid) = CreateControllerWithGrid(
                connector,
                metadataStore,
                dialogService,
                () => "Sheet1",
                aiClient);
            grid.SetCell("Sheet1", 3, 1, "ID");
            grid.SetCell("Sheet1", 3, 2, "项目负责人");

            InvokeRefresh(controller);
            InvokeExecuteAiColumnMapping(controller);

            Assert.Equal(1, dialogService.AiColumnMappingProgressRunCount);
            Assert.Equal(dialogService.LastProgressCancellationToken, aiClient.LastCancellationToken);
            Assert.Single(dialogService.AiColumnMappingPreviews);
            Assert.Empty(dialogService.ErrorMessages);
        }

        [Fact]
        public void ExecuteAiColumnMappingDoesNotSaveWhenProgressDialogCancelsOperation()
        {
            var connector = new FakeSystemConnector
            {
                FieldMappingDefinition = BuildAiMappingDefinition(),
            };
            var metadataStore = new FakeWorksheetMetadataStore();
            var dialogService = new FakeDialogService
            {
                AiColumnMappingConfirmResult = true,
                CancelAiColumnMappingProgress = true,
            };
            var aiClient = new FakeAiColumnMappingClient
            {
                Response = new AiColumnMappingResponse
                {
                    Mappings = new[]
                    {
                        new AiColumnMappingSuggestion
                        {
                            ExcelColumn = 2,
                            ActualL1 = "项目负责人",
                            TargetHeaderId = "owner_name",
                            TargetApiFieldKey = "owner_name",
                            Confidence = 0.91,
                        },
                    },
                },
            };
            metadataStore.Bindings["Sheet1"] = CreateAiMappingBinding();
            metadataStore.FieldMappings["Sheet1"] = BuildAiMappings("Sheet1");
            var (controller, grid) = CreateControllerWithGrid(
                connector,
                metadataStore,
                dialogService,
                () => "Sheet1",
                aiClient);
            grid.SetCell("Sheet1", 3, 1, "ID");
            grid.SetCell("Sheet1", 3, 2, "项目负责人");

            InvokeRefresh(controller);
            InvokeExecuteAiColumnMapping(controller);

            Assert.Equal(1, dialogService.AiColumnMappingProgressRunCount);
            Assert.Empty(dialogService.AiColumnMappingPreviews);
            Assert.Empty(metadataStore.LastSavedFieldMappings);
            Assert.Empty(dialogService.ErrorMessages);
        }

        [Fact]
        public void ExecuteFullDownloadSkipsConfirmationAndWriteWhenNoRowsMatch()
        {
            var connector = new FakeSystemConnector
            {
                FieldMappingDefinition = BuildAiMappingDefinition(),
                FindResult = Array.Empty<IDictionary<string, object>>(),
            };
            var metadataStore = new FakeWorksheetMetadataStore();
            var dialogService = new FakeDialogService();
            metadataStore.Bindings["Sheet1"] = CreateAiMappingBinding();
            metadataStore.FieldMappings["Sheet1"] = BuildAiMappings("Sheet1");
            var (controller, grid) = CreateControllerWithGrid(
                connector,
                metadataStore,
                dialogService,
                () => "Sheet1",
                aiClient: null);
            grid.SetCell("Sheet1", 1, 1, "existing header");

            InvokeRefresh(controller);
            InvokeExecuteFullDownload(controller);

            Assert.Equal(0, dialogService.ConfirmDownloadCallCount);
            Assert.Contains(dialogService.InfoMessages, message => message.IndexOf("query result is empty", StringComparison.OrdinalIgnoreCase) >= 0);
            Assert.Equal("existing header", grid.GetCell("Sheet1", 1, 1));
            Assert.Equal(0, grid.SetCellTextCallCount);
            Assert.Equal(0, grid.ClearRangeCallCount);
            Assert.Empty(dialogService.ErrorMessages);
        }

        [Fact]
        public void ExecutePartialDownloadSkipsConfirmationAndWriteWhenNoRowsMatch()
        {
            var connector = new FakeSystemConnector
            {
                FieldMappingDefinition = BuildAiMappingDefinition(),
                FindResult = Array.Empty<IDictionary<string, object>>(),
            };
            var metadataStore = new FakeWorksheetMetadataStore();
            var dialogService = new FakeDialogService();
            var selectionReader = new FakeWorksheetSelectionReader
            {
                Cells = new[]
                {
                    new SelectedVisibleCell { Row = 4, Column = 2 },
                },
            };
            metadataStore.Bindings["Sheet1"] = CreateAiMappingBinding();
            metadataStore.FieldMappings["Sheet1"] = BuildAiMappings("Sheet1");
            var (controller, grid) = CreateControllerWithGrid(
                connector,
                metadataStore,
                dialogService,
                () => "Sheet1",
                aiClient: null,
                selectionReader);
            grid.SetCell("Sheet1", 3, 1, "ID");
            grid.SetCell("Sheet1", 3, 2, "负责人");
            grid.SetCell("Sheet1", 4, 1, "row-1");
            grid.SetCell("Sheet1", 4, 2, "old owner");

            InvokeRefresh(controller);
            InvokeExecutePartialDownload(controller);

            Assert.Equal(0, dialogService.ConfirmDownloadCallCount);
            Assert.Equal("performance", connector.LastFindProjectId);
            Assert.Equal(new[] { "row-1" }, connector.LastFindRowIds);
            Assert.Contains(dialogService.InfoMessages, message => message.IndexOf("query result is empty", StringComparison.OrdinalIgnoreCase) >= 0);
            Assert.Equal("old owner", grid.GetCell("Sheet1", 4, 2));
            Assert.Equal(0, grid.WriteRangeValuesCallCount);
            Assert.Empty(dialogService.ErrorMessages);
        }

        private static object CreateController(
            FakeSystemConnector connector,
            string activeSheetName,
            IAnalyticsService analyticsService = null)
        {
            return CreateController(
                connector,
                new FakeWorksheetMetadataStore(),
                new FakeDialogService { ConfirmProjectLayoutWithSuggestedBinding = true },
                () => activeSheetName,
                authenticationLoginAction: null,
                analyticsService: analyticsService);
        }

        private static object CreateController(
            FakeSystemConnector connector,
            FakeWorksheetMetadataStore metadataStore,
            FakeDialogService dialogService,
            Func<string> sheetNameProvider,
            IAnalyticsService analyticsService = null)
        {
            return CreateController(connector, metadataStore, dialogService, sheetNameProvider, null, analyticsService: analyticsService);
        }

        private static object CreateController(
            FakeSystemConnector connector,
            FakeWorksheetMetadataStore metadataStore,
            FakeDialogService dialogService,
            Func<string> sheetNameProvider,
            Action authenticationLoginAction,
            Action authenticationRequiredAction = null,
            IAnalyticsService analyticsService = null)
        {
            var addInAssembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var controllerType = addInAssembly.GetType("OfficeAgent.ExcelAddIn.RibbonSyncController", throwOnError: true);
            var (executionService, _) = CreateExecutionService(addInAssembly, connector, metadataStore);
            var dialogInterface = addInAssembly.GetType("OfficeAgent.ExcelAddIn.Dialogs.IRibbonSyncDialogService", throwOnError: true);
            var syncService = new WorksheetSyncService(
                new SystemConnectorRegistry(new ISystemConnector[] { connector }),
                metadataStore,
                new WorksheetChangeTracker(),
                new SyncOperationPreviewFactory());

            var ctorTypes = analyticsService != null
                ? new[]
                {
                    typeof(IWorksheetMetadataStore),
                    typeof(WorksheetSyncService),
                    typeof(Func<string>),
                    executionService.GetType(),
                    dialogInterface,
                    typeof(Action),
                    typeof(Action),
                    typeof(IAnalyticsService),
                }
                : authenticationLoginAction == null && authenticationRequiredAction == null
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

            if (analyticsService != null)
            {
                return ctor.Invoke(new object[] { metadataStore, syncService, sheetNameProvider, executionService, dialogService.GetTransparentProxy(), authenticationLoginAction, authenticationRequiredAction, analyticsService });
            }

            return authenticationLoginAction == null && authenticationRequiredAction == null
                ? ctor.Invoke(new object[] { metadataStore, syncService, sheetNameProvider, executionService, dialogService.GetTransparentProxy() })
                : ctor.Invoke(new object[] { metadataStore, syncService, sheetNameProvider, executionService, dialogService.GetTransparentProxy(), authenticationLoginAction, authenticationRequiredAction });
        }

        private static (object Controller, FakeWorksheetGridAdapter Grid) CreateControllerWithGrid(
            FakeSystemConnector connector,
            FakeWorksheetMetadataStore metadataStore,
            FakeDialogService dialogService,
            Func<string> sheetNameProvider,
            IAiColumnMappingClient aiClient,
            FakeWorksheetSelectionReader selectionReader = null,
            IAnalyticsService analyticsService = null)
        {
            var addInAssembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var controllerType = addInAssembly.GetType("OfficeAgent.ExcelAddIn.RibbonSyncController", throwOnError: true);
            var (executionService, grid) = CreateExecutionService(addInAssembly, connector, metadataStore, aiClient, selectionReader);
            var dialogInterface = addInAssembly.GetType("OfficeAgent.ExcelAddIn.Dialogs.IRibbonSyncDialogService", throwOnError: true);
            var syncService = new WorksheetSyncService(
                new SystemConnectorRegistry(new ISystemConnector[] { connector }),
                metadataStore,
                new WorksheetChangeTracker(),
                new SyncOperationPreviewFactory());
            var ctorTypes = analyticsService == null
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
                    typeof(Action),
                    typeof(IAnalyticsService),
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

            var args = analyticsService == null
                ? new object[] { metadataStore, syncService, sheetNameProvider, executionService, dialogService.GetTransparentProxy() }
                : new object[] { metadataStore, syncService, sheetNameProvider, executionService, dialogService.GetTransparentProxy(), null, null, analyticsService };

            return (ctor.Invoke(args), grid);
        }

        private static (object Service, FakeWorksheetGridAdapter Grid) CreateExecutionService(
            Assembly addInAssembly,
            FakeSystemConnector connector,
            FakeWorksheetMetadataStore metadataStore,
            IAiColumnMappingClient aiClient = null,
            FakeWorksheetSelectionReader selectionReader = null)
        {
            var serviceType = addInAssembly.GetType("OfficeAgent.ExcelAddIn.WorksheetSyncExecutionService", throwOnError: true);
            var gridInterface = addInAssembly.GetType("OfficeAgent.ExcelAddIn.Excel.IWorksheetGridAdapter", throwOnError: true);
            var grid = new FakeWorksheetGridAdapter(gridInterface);
            var syncService = new WorksheetSyncService(
                new SystemConnectorRegistry(new ISystemConnector[] { connector }),
                metadataStore,
                new WorksheetChangeTracker(),
                new SyncOperationPreviewFactory());

            if (connector is IBusinessExportTemplateConnector)
            {
                var fullCtor = FindConstructor(
                    serviceType,
                    typeof(WorksheetSyncService),
                    typeof(IWorksheetMetadataStore),
                    typeof(IWorksheetSelectionReader),
                    gridInterface,
                    typeof(SyncOperationPreviewFactory),
                    "OfficeAgent.ExcelAddIn.Excel.IWorksheetChangeLogStore",
                    "OfficeAgent.ExcelAddIn.Excel.WorksheetPendingEditTracker",
                    typeof(IAiColumnMappingClient),
                    "OfficeAgent.ExcelAddIn.Excel.IBusinessWorkbookImporter");

                if (fullCtor == null)
                {
                    throw new InvalidOperationException("WorksheetSyncExecutionService template import constructor was not found.");
                }

                var businessWorkbookImporter = new FakeBusinessWorkbookImporter(
                    addInAssembly.GetType("OfficeAgent.ExcelAddIn.Excel.IBusinessWorkbookImporter", throwOnError: true));

                return (fullCtor.Invoke(new object[]
                {
                    syncService,
                    metadataStore,
                    selectionReader ?? new FakeWorksheetSelectionReader(),
                    grid.GetTransparentProxy(),
                    new SyncOperationPreviewFactory(),
                    null,
                    null,
                    aiClient ?? new FakeAiColumnMappingClient(),
                    businessWorkbookImporter.GetTransparentProxy(),
                }), grid);
            }

            var ctorTypes = aiClient == null
                ? new[]
                {
                    typeof(WorksheetSyncService),
                    typeof(IWorksheetMetadataStore),
                    typeof(IWorksheetSelectionReader),
                    gridInterface,
                    typeof(SyncOperationPreviewFactory),
                }
                : new[]
                {
                    typeof(WorksheetSyncService),
                    typeof(IWorksheetMetadataStore),
                    typeof(IWorksheetSelectionReader),
                    gridInterface,
                    typeof(SyncOperationPreviewFactory),
                    typeof(IAiColumnMappingClient),
                };
            var ctor = serviceType.GetConstructor(
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                binder: null,
                types: ctorTypes,
                modifiers: null);

            if (ctor == null)
            {
                throw new InvalidOperationException("WorksheetSyncExecutionService constructor was not found.");
            }

            var args = aiClient == null
                ? new object[]
                {
                    syncService,
                    metadataStore,
                    selectionReader ?? new FakeWorksheetSelectionReader(),
                    grid.GetTransparentProxy(),
                    new SyncOperationPreviewFactory(),
                }
                : new object[]
                {
                    syncService,
                    metadataStore,
                    selectionReader ?? new FakeWorksheetSelectionReader(),
                    grid.GetTransparentProxy(),
                    new SyncOperationPreviewFactory(),
                    aiClient,
                };

            return (ctor.Invoke(args), grid);
        }

        private static ConstructorInfo FindConstructor(Type type, params object[] expectedTypes)
        {
            return type.GetConstructors(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
                .FirstOrDefault(constructor =>
                {
                    var parameters = constructor.GetParameters();
                    if (parameters.Length != expectedTypes.Length)
                    {
                        return false;
                    }

                    for (var index = 0; index < expectedTypes.Length; index++)
                    {
                        var parameterType = parameters[index].ParameterType;
                        if (expectedTypes[index] is Type expectedType)
                        {
                            if (!parameterType.IsAssignableFrom(expectedType) &&
                                !expectedType.IsAssignableFrom(parameterType) &&
                                !string.Equals(parameterType.FullName, expectedType.FullName, StringComparison.Ordinal))
                            {
                                return false;
                            }

                            continue;
                        }

                        if (!string.Equals(parameterType.FullName, Convert.ToString(expectedTypes[index]), StringComparison.Ordinal))
                        {
                            return false;
                        }
                    }

                    return true;
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

        private static List<OfficeAgentLogEntry> CaptureLogEntries(Action action)
        {
            var capture = CaptureLogEntriesAllowingFailure(action);
            if (capture.Failure != null)
            {
                throw capture.Failure;
            }

            return capture.Entries;
        }

        private static LogCaptureResult CaptureLogEntriesAllowingFailure(Action action)
        {
            var entries = new List<OfficeAgentLogEntry>();
            OfficeAgentLog.Configure(entries.Add);

            try
            {
                action();
                return new LogCaptureResult(entries, null);
            }
            catch (Exception ex)
            {
                return new LogCaptureResult(entries, ex);
            }
            finally
            {
                OfficeAgentLog.Reset();
            }
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

        private static void InvokeExecuteFullDownload(object controller)
        {
            var method = controller.GetType().GetMethod(
                "ExecuteFullDownload",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

            if (method == null)
            {
                throw new InvalidOperationException("RibbonSyncController.ExecuteFullDownload() was not found.");
            }

            method.Invoke(controller, null);
        }

        private static void InvokeExecutePartialDownload(object controller)
        {
            var method = controller.GetType().GetMethod(
                "ExecutePartialDownload",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

            if (method == null)
            {
                throw new InvalidOperationException("RibbonSyncController.ExecutePartialDownload() was not found.");
            }

            method.Invoke(controller, null);
        }

        private static void InvokeExecuteAiColumnMapping(object controller)
        {
            var method = controller.GetType().GetMethod(
                "ExecuteAiColumnMapping",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

            if (method == null)
            {
                throw new InvalidOperationException("RibbonSyncController.ExecuteAiColumnMapping() was not found.");
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

        private static string ResolveRepositoryPath(params string[] segments)
        {
            return Path.GetFullPath(Path.Combine(new[]
            {
                AppContext.BaseDirectory,
                "..",
                "..",
                "..",
                "..",
                "..",
            }.Concat(segments).ToArray()));
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
            var stringsType = addInAssembly.GetType("OfficeAgent.ExcelAddIn.Localization.HostLocalizedStrings", throwOnError: true);
            var strings = stringsType
                .GetMethod("ForLocale", BindingFlags.Static | BindingFlags.Public)
                .Invoke(null, new object[] { "zh" });
            var method = controllerType.GetMethod(
                "BuildUploadPreviewInfoMessage",
                BindingFlags.Static | BindingFlags.NonPublic);
            if (method == null)
            {
                throw new InvalidOperationException("RibbonSyncController.BuildUploadPreviewInfoMessage was not found.");
            }

            return (string)method.Invoke(null, new[] { strings, operationName, preview });
        }

        private static SheetBinding CreateAiMappingBinding()
        {
            return new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 3,
                HeaderRowCount = 1,
                DataStartRow = 4,
            };
        }

        private static SheetBinding CreateInitializedBinding(string systemKey)
        {
            return new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = systemKey,
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 1,
                HeaderRowCount = 2,
                DataStartRow = 3,
            };
        }

        private static void AssertTemplateImportProperties(
            AnalyticsEvent analyticsEvent,
            string templateId,
            string templateName,
            bool isBlankSheet)
        {
            Assert.Equal("performance", analyticsEvent.Properties["projectId"]);
            Assert.Equal("绩效项目", analyticsEvent.Properties["projectName"]);
            Assert.Equal("Sheet1", analyticsEvent.Properties["sheetName"]);
            Assert.Equal(templateId, analyticsEvent.Properties["templateId"]);
            Assert.Equal(templateName, analyticsEvent.Properties["templateName"]);
            Assert.Equal(isBlankSheet, analyticsEvent.Properties["isBlankSheet"]);
            Assert.True(analyticsEvent.Properties.ContainsKey("durationMs"));
        }

        private static FieldMappingTableDefinition BuildAiMappingDefinition()
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
                },
            };
        }

        private static SheetFieldMappingRow[] BuildAiMappings(string sheetName)
        {
            return new[]
            {
                CreateAiMappingRow(sheetName, "row_id", "ID", true),
                CreateAiMappingRow(sheetName, "owner_name", "负责人", false),
            };
        }

        private static SheetFieldMappingRow CreateAiMappingRow(
            string sheetName,
            string apiFieldKey,
            string headerText,
            bool isIdColumn)
        {
            return new SheetFieldMappingRow
            {
                SheetName = sheetName,
                Values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["HeaderId"] = apiFieldKey,
                    ["HeaderType"] = "single",
                    ["ApiFieldKey"] = apiFieldKey,
                    ["IsIdColumn"] = isIdColumn ? "true" : "false",
                    ["DefaultL1"] = headerText,
                    ["CurrentL1"] = headerText,
                    ["DefaultL2"] = string.Empty,
                    ["CurrentL2"] = string.Empty,
                },
            };
        }

        private class FakeSystemConnector : ISystemConnector
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

            public IReadOnlyList<IDictionary<string, object>> FindResult { get; set; } = Array.Empty<IDictionary<string, object>>();

            public string LastBuildFieldMappingSeedProjectId { get; private set; }

            public string LastFindProjectId { get; private set; }

            public IReadOnlyList<string> LastFindRowIds { get; private set; } = Array.Empty<string>();

            public IReadOnlyList<string> LastFindFieldKeys { get; private set; } = Array.Empty<string>();

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

            public IReadOnlyList<IDictionary<string, object>> Find(
                string projectId,
                IReadOnlyList<string> rowIds,
                IReadOnlyList<string> fieldKeys)
            {
                LastFindProjectId = projectId;
                LastFindRowIds = rowIds?.ToArray() ?? Array.Empty<string>();
                LastFindFieldKeys = fieldKeys?.ToArray() ?? Array.Empty<string>();
                return FindResult;
            }

            public void BatchSave(string projectId, IReadOnlyList<CellChange> changes)
            {
                throw new NotSupportedException();
            }
        }

        private sealed class FakeBusinessTemplateConnector : FakeSystemConnector, IBusinessExportTemplateConnector
        {
            public IReadOnlyList<BusinessExportTemplateOption> TemplateOptions { get; set; } = new[]
            {
                new BusinessExportTemplateOption
                {
                    TemplateId = "standard",
                    TemplateName = "标准作业表",
                },
            };

            public Exception ExportException { get; set; }

            public string LastGetBusinessExportTemplatesProjectId { get; private set; }

            public string LastExportProjectId { get; private set; }

            public string LastExportTemplateId { get; private set; }

            public IReadOnlyList<BusinessExportTemplateOption> GetBusinessExportTemplates(string projectId)
            {
                LastGetBusinessExportTemplatesProjectId = projectId;
                return TemplateOptions;
            }

            public Task<BusinessExportWorkbook> ExportBusinessWorkbookAsync(
                string projectId,
                string templateId,
                CancellationToken cancellationToken)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (ExportException != null)
                {
                    throw ExportException;
                }

                LastExportProjectId = projectId;
                LastExportTemplateId = templateId;
                return Task.FromResult(new BusinessExportWorkbook
                {
                    Content = new byte[] { 0x50, 0x4B, 0x03, 0x04 },
                });
            }
        }

        private sealed class FakeBusinessWorkbookImporter : RealProxy
        {
            public FakeBusinessWorkbookImporter(Type interfaceType)
                : base(interfaceType)
            {
            }

            public bool IsBlank { get; set; } = true;

            public int IsBlankCallCount { get; private set; }

            public int ImportCallCount { get; private set; }

            public string ImportedTargetSheetName { get; private set; }

            public override IMessage Invoke(IMessage msg)
            {
                var call = (IMethodCallMessage)msg;
                switch (call.MethodName)
                {
                    case "IsWorkSheetContentBlank":
                        IsBlankCallCount++;
                        return new ReturnMessage(IsBlank, null, 0, call.LogicalCallContext, call);
                    case "EnsureCanWriteToWorkSheet":
                        return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
                    case "ImportBusinessDataSheet":
                        ImportCallCount++;
                        ImportedTargetSheetName = (string)call.InArgs[1];
                        return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
                    case "ActivateWorkSheetAtA1":
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
                cancellationToken.ThrowIfCancellationRequested();
                return System.Threading.Tasks.Task.FromResult(Response);
            }
        }

        private sealed class RecordingAnalyticsService : IAnalyticsService
        {
            public List<AnalyticsEvent> Events { get; } = new List<AnalyticsEvent>();

            public void Track(AnalyticsEvent analyticsEvent)
            {
                Events.Add(analyticsEvent);
            }

            public void Track(
                string eventName,
                string source,
                IDictionary<string, object> properties = null,
                IDictionary<string, object> businessContext = null,
                AnalyticsError error = null)
            {
                Track(new AnalyticsEvent
                {
                    EventName = eventName,
                    Source = source,
                    Properties = properties,
                    BusinessContext = businessContext,
                    Error = error,
                });
            }
        }

        private sealed class FakeWorksheetMetadataStore : IWorksheetMetadataStore
        {
            public Dictionary<string, SheetBinding> Bindings { get; } = new Dictionary<string, SheetBinding>(StringComparer.OrdinalIgnoreCase);

            public Dictionary<string, SheetFieldMappingRow[]> FieldMappings { get; } = new Dictionary<string, SheetFieldMappingRow[]>(StringComparer.OrdinalIgnoreCase);

            public int LoadBindingCallCount { get; private set; }

            public SheetBinding LastSavedBinding { get; private set; }

            public string LastRefreshedPresentationSheetName { get; private set; }

            public bool LastRefreshedPresentationHideTemplateBindingRows { get; private set; }

            public SheetFieldMappingRow[] LastSavedFieldMappings { get; private set; } = Array.Empty<SheetFieldMappingRow>();

            public Exception SaveBindingException { get; set; }

            public void SaveBinding(SheetBinding binding)
            {
                if (SaveBindingException != null)
                {
                    throw SaveBindingException;
                }

                LastSavedBinding = binding;
                Bindings[binding.SheetName] = binding;
            }

            public void RefreshMetadataPresentation(string sheetName, bool hideTemplateBindingRows = false)
            {
                LastRefreshedPresentationSheetName = sheetName;
                LastRefreshedPresentationHideTemplateBindingRows = hideTemplateBindingRows;
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

            public List<AiColumnMappingPreview> AiColumnMappingPreviews { get; } = new List<AiColumnMappingPreview>();

            public List<InitializeSheetDialogRequest> InitializeSheetRequests { get; } = new List<InitializeSheetDialogRequest>();

            public List<InitializeSheetTemplateLoadResult> InitializeSheetTemplateLoadResults { get; } = new List<InitializeSheetTemplateLoadResult>();

            public List<string> InitializeTemplateImportProgressCalls { get; } = new List<string>();

            public int ConfirmDownloadCallCount { get; private set; }

            public int AiColumnMappingProgressRunCount { get; private set; }

            public int InitializeTemplateImportProgressRunCount { get; private set; }

            public CancellationToken LastProgressCancellationToken { get; private set; }

            public SheetBinding NextProjectLayoutBinding { get; set; }

            public InitializeSheetDialogResult NextInitializeSheetResult { get; set; }

            public bool CancelInitializeSheetDialog { get; set; }

            public bool ConfirmProjectLayoutWithSuggestedBinding { get; set; }

            public bool AuthenticationRequiredResult { get; set; }

            public bool AiColumnMappingConfirmResult { get; set; }

            public bool CancelAiColumnMappingProgress { get; set; }

            public bool CancelInitializeTemplateImportProgress { get; set; }

            public Action<AiColumnMappingPreview> OnConfirmAiColumnMapping { get; set; }

            public override IMessage Invoke(IMessage msg)
            {
                var call = (IMethodCallMessage)msg;
                switch (call.MethodName)
                {
                    case "ConfirmDownload":
                        ConfirmDownloadCallCount++;
                        return new ReturnMessage(true, null, 0, call.LogicalCallContext, call);
                    case "ConfirmUpload":
                        return new ReturnMessage(true, null, 0, call.LogicalCallContext, call);
                    case "ConfirmAiColumnMapping":
                        var preview = (AiColumnMappingPreview)call.InArgs[0];
                        AiColumnMappingPreviews.Add(preview);
                        OnConfirmAiColumnMapping?.Invoke(preview);
                        return new ReturnMessage(AiColumnMappingConfirmResult, null, 0, call.LogicalCallContext, call);
                    case "RunAiColumnMappingWithProgress":
                        AiColumnMappingProgressRunCount++;
                        var operation = (Func<CancellationToken, Task<AiColumnMappingPreview>>)call.InArgs[0];
                        using (var cancellationTokenSource = new CancellationTokenSource())
                        {
                            if (CancelAiColumnMappingProgress)
                            {
                                cancellationTokenSource.Cancel();
                            }

                            LastProgressCancellationToken = cancellationTokenSource.Token;
                            try
                            {
                                return new ReturnMessage(operation(cancellationTokenSource.Token).GetAwaiter().GetResult(), null, 0, call.LogicalCallContext, call);
                            }
                            catch (OperationCanceledException)
                            {
                                return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
                            }
                        }
                    case "ShowProjectLayoutDialog":
                        ProjectLayoutPrompts.Add(CloneBinding((SheetBinding)call.InArgs[0]));
                        var layoutBinding = ConfirmProjectLayoutWithSuggestedBinding
                            ? (SheetBinding)call.InArgs[0]
                            : NextProjectLayoutBinding;
                        return new ReturnMessage(CloneBinding(layoutBinding), null, 0, call.LogicalCallContext, call);
                    case "ShowInitializeSheetDialog":
                        InitializeSheetRequests.Add(CloneInitializeSheetRequest(call.InArgs[0]));
                        InitializeSheetTemplateLoadResults.Add(InvokeLoadTemplates(call.InArgs[1]));
                        return new ReturnMessage(
                            CancelInitializeSheetDialog
                                ? null
                                : CreateInitializeSheetResult(call.MethodBase, NextInitializeSheetResult),
                            null,
                            0,
                            call.LogicalCallContext,
                            call);
                    case "RunInitializeSheetTemplateImportWithProgress":
                        return RunInitializeSheetTemplateImportWithProgress(call);
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

            private IMessage RunInitializeSheetTemplateImportWithProgress(IMethodCallMessage call)
            {
                InitializeTemplateImportProgressRunCount++;
                var operation = call.InArgs[0];
                using (var cancellationTokenSource = new CancellationTokenSource())
                {
                    if (CancelInitializeTemplateImportProgress)
                    {
                        cancellationTokenSource.Cancel();
                    }

                    LastProgressCancellationToken = cancellationTokenSource.Token;
                    try
                    {
                        var progressInterface = call.MethodBase
                            .GetParameters()[0]
                            .ParameterType
                            .GetGenericArguments()[0];
                        var progress = new FakeInitializeSheetImportProgress(
                            progressInterface,
                            InitializeTemplateImportProgressCalls).GetTransparentProxy();
                        var task = (Task)operation.GetType()
                            .GetMethod("Invoke")
                            .Invoke(operation, new[] { progress, cancellationTokenSource.Token });
                        task.GetAwaiter().GetResult();
                        return new ReturnMessage(!cancellationTokenSource.IsCancellationRequested, null, 0, call.LogicalCallContext, call);
                    }
                    catch (TargetInvocationException ex) when (ex.InnerException != null)
                    {
                        if (ex.InnerException is OperationCanceledException && cancellationTokenSource.IsCancellationRequested)
                        {
                            return new ReturnMessage(false, null, 0, call.LogicalCallContext, call);
                        }

                        ExceptionDispatchInfo.Capture(ex.InnerException).Throw();
                        throw;
                    }
                    catch (OperationCanceledException)
                    {
                        if (cancellationTokenSource.IsCancellationRequested)
                        {
                            return new ReturnMessage(false, null, 0, call.LogicalCallContext, call);
                        }

                        throw;
                    }
                }
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

            private static InitializeSheetDialogRequest CloneInitializeSheetRequest(object request)
            {
                if (request == null)
                {
                    return null;
                }

                return new InitializeSheetDialogRequest
                {
                    ProjectDisplayName = ReadProperty<string>(request, "ProjectDisplayName"),
                    IsBlankSheet = ReadProperty<bool>(request, "IsBlankSheet"),
                    SupportsTemplateImport = ReadProperty<bool>(request, "SupportsTemplateImport"),
                };
            }

            private static InitializeSheetTemplateLoadResult InvokeLoadTemplates(object loadTemplates)
            {
                var loadResult = loadTemplates.GetType()
                    .GetMethod("Invoke")
                    .Invoke(loadTemplates, null);
                return InitializeSheetTemplateLoadResult.Success(
                    ((IEnumerable<BusinessExportTemplateOption>)ReadProperty<object>(loadResult, "Templates"))
                    .Select(CloneTemplate));
            }

            private static object CreateInitializeSheetResult(MethodBase method, InitializeSheetDialogResult result)
            {
                var resultType = ((MethodInfo)method).ReturnType;
                var modeType = resultType.GetProperty("Mode").PropertyType;
                var templateType = resultType.GetProperty("SelectedTemplate").PropertyType;
                var target = Activator.CreateInstance(resultType);
                var selected = result ?? new InitializeSheetDialogResult
                {
                    Mode = InitializeSheetMode.ConfigOnly,
                };

                resultType.GetProperty("Mode").SetValue(
                    target,
                    Enum.Parse(modeType, selected.Mode.ToString()));
                resultType.GetProperty("SelectedTemplate").SetValue(
                    target,
                    CloneTemplate(templateType, selected.SelectedTemplate));

                return target;
            }

            private static T ReadProperty<T>(object target, string propertyName)
            {
                var value = target.GetType().GetProperty(propertyName).GetValue(target);
                if (value == null)
                {
                    return default(T);
                }

                return (T)value;
            }

            private static object CloneTemplate(Type templateType, BusinessExportTemplateOption template)
            {
                if (template == null)
                {
                    return null;
                }

                var target = Activator.CreateInstance(templateType);
                templateType.GetProperty("TemplateId").SetValue(target, template.TemplateId);
                templateType.GetProperty("TemplateName").SetValue(target, template.TemplateName);
                return target;
            }

            private static BusinessExportTemplateOption CloneTemplate(BusinessExportTemplateOption template)
            {
                if (template == null)
                {
                    return null;
                }

                return new BusinessExportTemplateOption
                {
                    TemplateId = template.TemplateId,
                    TemplateName = template.TemplateName,
                };
            }
        }

        private sealed class FakeInitializeSheetImportProgress : RealProxy
        {
            private readonly List<string> calls;

            public FakeInitializeSheetImportProgress(Type interfaceType, List<string> calls)
                : base(interfaceType)
            {
                this.calls = calls ?? throw new ArgumentNullException(nameof(calls));
            }

            public override IMessage Invoke(IMessage msg)
            {
                var call = (IMethodCallMessage)msg;
                switch (call.MethodName)
                {
                    case "SetDownloading":
                        calls.Add("downloading");
                        return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
                    case "SetImporting":
                        calls.Add("importing");
                        return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
                    case "SetWritingConfiguration":
                        calls.Add("writingConfiguration");
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

        private sealed class LogCaptureResult
        {
            public LogCaptureResult(List<OfficeAgentLogEntry> entries, Exception failure)
            {
                Entries = entries;
                Failure = failure;
            }

            public List<OfficeAgentLogEntry> Entries { get; }

            public Exception Failure { get; }
        }

        private sealed class FakeWorksheetSelectionReader : IWorksheetSelectionReader
        {
            public IReadOnlyList<SelectedVisibleCell> Cells { get; set; } = Array.Empty<SelectedVisibleCell>();

            public WorksheetSelectionSnapshot SelectionSnapshot { get; set; } = new WorksheetSelectionSnapshot();

            public IReadOnlyList<SelectedVisibleCell> ReadVisibleSelection()
            {
                return Cells;
            }

            public WorksheetSelectionSnapshot ReadSelectionSnapshot()
            {
                return SelectionSnapshot;
            }
        }

        private sealed class FakeWorksheetGridAdapter : RealProxy
        {
            private readonly Dictionary<(string Sheet, int Row, int Column), string> cells =
                new Dictionary<(string Sheet, int Row, int Column), string>();

            public FakeWorksheetGridAdapter(Type interfaceType)
                : base(interfaceType)
            {
            }

            public void SetCell(string sheetName, int row, int column, string value)
            {
                cells[(sheetName, row, column)] = value ?? string.Empty;
            }

            public int SetCellTextCallCount { get; private set; }

            public int ClearRangeCallCount { get; private set; }

            public int WriteRangeValuesCallCount { get; private set; }

            public string GetCell(string sheetName, int row, int column)
            {
                cells.TryGetValue((sheetName, row, column), out var value);
                return value ?? string.Empty;
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
                        return HandleGetLastUsedRow(call);
                    case "SetCellText":
                        SetCellTextCallCount++;
                        return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
                    case "ClearRange":
                        ClearRangeCallCount++;
                        return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
                    case "WriteRangeValues":
                        WriteRangeValuesCallCount++;
                        return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
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

            private IMessage HandleGetLastUsedRow(IMethodCallMessage call)
            {
                var sheetName = (string)call.InArgs[0];
                var lastRow = cells.Keys
                    .Where(key => string.Equals(key.Sheet, sheetName, StringComparison.OrdinalIgnoreCase))
                    .Select(key => key.Row)
                    .DefaultIfEmpty(0)
                    .Max();
                return new ReturnMessage(lastRow, null, 0, call.LogicalCallContext, call);
            }
        }
    }
}
