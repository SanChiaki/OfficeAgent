using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeAgent.Core;
using OfficeAgent.Core.Analytics;
using OfficeAgent.Core.Diagnostics;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
using OfficeAgent.Core.Sync;
using OfficeAgent.ExcelAddIn.Dialogs;
using OfficeAgent.ExcelAddIn.Localization;

namespace OfficeAgent.ExcelAddIn
{
    internal sealed class RibbonSyncController
    {
        private readonly IWorksheetMetadataStore metadataStore;
        private readonly WorksheetSyncService worksheetSyncService;
        private readonly Func<string> activeSheetNameProvider;
        private readonly WorksheetSyncExecutionService executionService;
        private readonly IRibbonSyncDialogService dialogService;
        private readonly Action authenticationLoginAction;
        private readonly IAnalyticsService analyticsService;
        private string lastRefreshedSheetName;

        public RibbonSyncController(
            IWorksheetMetadataStore metadataStore,
            WorksheetSyncService worksheetSyncService,
            Func<string> activeSheetNameProvider)
            : this(
                metadataStore,
                worksheetSyncService,
                activeSheetNameProvider,
                executionService: null,
                new RibbonSyncDialogService(),
                authenticationLoginAction: null)
        {
        }

        internal RibbonSyncController(
            IWorksheetMetadataStore metadataStore,
            WorksheetSyncService worksheetSyncService,
            Func<string> activeSheetNameProvider,
            WorksheetSyncExecutionService executionService)
            : this(
                metadataStore,
                worksheetSyncService,
                activeSheetNameProvider,
                executionService,
                new RibbonSyncDialogService(),
                authenticationLoginAction: null)
        {
        }

        internal RibbonSyncController(
            IWorksheetMetadataStore metadataStore,
            WorksheetSyncService worksheetSyncService,
            Func<string> activeSheetNameProvider,
            WorksheetSyncExecutionService executionService,
            IRibbonSyncDialogService dialogService)
            : this(
                metadataStore,
                worksheetSyncService,
                activeSheetNameProvider,
                executionService,
                dialogService,
                authenticationLoginAction: null)
        {
        }

        internal RibbonSyncController(
            IWorksheetMetadataStore metadataStore,
            WorksheetSyncService worksheetSyncService,
            Func<string> activeSheetNameProvider,
            WorksheetSyncExecutionService executionService,
            IRibbonSyncDialogService dialogService,
            Action authenticationLoginAction)
            : this(
                metadataStore,
                worksheetSyncService,
                activeSheetNameProvider,
                executionService,
                dialogService,
                authenticationLoginAction,
                analyticsService: null)
        {
        }

        internal RibbonSyncController(
            IWorksheetMetadataStore metadataStore,
            WorksheetSyncService worksheetSyncService,
            Func<string> activeSheetNameProvider,
            WorksheetSyncExecutionService executionService,
            IRibbonSyncDialogService dialogService,
            Action authenticationLoginAction,
            IAnalyticsService analyticsService = null)
        {
            this.metadataStore = metadataStore ?? throw new ArgumentNullException(nameof(metadataStore));
            this.worksheetSyncService = worksheetSyncService ?? throw new ArgumentNullException(nameof(worksheetSyncService));
            this.activeSheetNameProvider = activeSheetNameProvider ?? throw new ArgumentNullException(nameof(activeSheetNameProvider));
            this.executionService = executionService;
            this.dialogService = dialogService ?? throw new ArgumentNullException(nameof(dialogService));
            this.authenticationLoginAction = authenticationLoginAction;
            this.analyticsService = analyticsService ?? NoopAnalyticsService.Instance;

            ActiveProjectDisplayName = GetStrings().ProjectDropDownPlaceholderText;
            ActiveProjectId = string.Empty;
            ActiveSystemKey = string.Empty;
        }

        public event EventHandler ActiveProjectChanged;

        public string ActiveProjectDisplayName { get; private set; }

        public string ActiveProjectId { get; private set; }

        public string ActiveSystemKey { get; private set; }

        internal SheetBinding ActiveBinding { get; private set; }

        public IReadOnlyList<ProjectOption> GetProjects()
        {
            return worksheetSyncService.GetProjects() ?? Array.Empty<ProjectOption>();
        }

        public void SelectProject(ProjectOption project)
        {
            if (project == null)
            {
                throw new ArgumentNullException(nameof(project));
            }

            var sheetName = GetRequiredSheetName();
            var existingBinding = TryLoadBinding(sheetName);

            OfficeAgentLog.Info(
                "ribbon_sync",
                "project.select.begin",
                "Project selection started.",
                BuildProjectSelectionDetails(sheetName, project, existingBinding));

            if (IsSameProject(existingBinding, project))
            {
                lastRefreshedSheetName = sheetName;
                ApplyBindingState(existingBinding);
                TrackRibbonEvent("ribbon.project.selected");
                OfficeAgentLog.Info(
                    "ribbon_sync",
                    "project.select.same_project",
                    "Selected project already matches the active worksheet binding.",
                    BuildProjectSelectionDetails(sheetName, project, existingBinding));
                return;
            }

            var suggestedBinding = worksheetSyncService.CreateBindingSeed(sheetName, project);
            OfficeAgentLog.Info(
                "ribbon_sync",
                "project.layout_dialog.show",
                "Showing project layout dialog.",
                BuildProjectSelectionDetails(sheetName, project, existingBinding, suggestedBinding));

            var confirmedBinding = dialogService.ShowProjectLayoutDialog(suggestedBinding);

            if (confirmedBinding == null)
            {
                OfficeAgentLog.Warn(
                    "ribbon_sync",
                    "project.layout_dialog.cancelled",
                    "Project layout dialog returned without confirmation.",
                    BuildProjectSelectionDetails(sheetName, project, existingBinding, suggestedBinding, includeRestoredProject: true));
                RestoreBindingState(existingBinding, sheetName);
                TrackRibbonEvent(
                    "ribbon.project_layout.canceled",
                    new Dictionary<string, object>(StringComparer.Ordinal)
                    {
                        ["targetSystemKey"] = project.SystemKey ?? string.Empty,
                        ["targetProjectId"] = project.ProjectId ?? string.Empty,
                        ["targetProjectName"] = project.DisplayName ?? string.Empty,
                    });
                return;
            }

            TrackRibbonEvent(
                "ribbon.project_layout.confirmed",
                new Dictionary<string, object>(StringComparer.Ordinal)
                {
                    ["targetSystemKey"] = confirmedBinding.SystemKey ?? string.Empty,
                    ["targetProjectId"] = confirmedBinding.ProjectId ?? string.Empty,
                    ["targetProjectName"] = confirmedBinding.ProjectName ?? string.Empty,
                    ["headerStartRow"] = confirmedBinding.HeaderStartRow,
                    ["headerRowCount"] = confirmedBinding.HeaderRowCount,
                    ["dataStartRow"] = confirmedBinding.DataStartRow,
                });

            OfficeAgentLog.Info(
                "ribbon_sync",
                "project.binding.save.begin",
                "Saving selected project binding.",
                BuildProjectSelectionDetails(sheetName, project, existingBinding, confirmedBinding));
            try
            {
                metadataStore.ClearFieldMappings(sheetName);
                metadataStore.SaveBinding(confirmedBinding);
            }
            catch (Exception ex)
            {
                OfficeAgentLog.Error(
                    "ribbon_sync",
                    "project.binding.save.failed",
                    "Failed to save selected project binding.",
                    ex,
                    BuildProjectSelectionDetails(sheetName, project, existingBinding, confirmedBinding));
                throw;
            }

            OfficeAgentLog.Info(
                "ribbon_sync",
                "project.binding.save.completed",
                "Selected project binding saved.",
                BuildProjectSelectionDetails(sheetName, project, existingBinding, confirmedBinding));
            lastRefreshedSheetName = sheetName;
            ApplyBindingState(confirmedBinding);
            TrackRibbonEvent("ribbon.project.selected");
        }

        public void RefreshActiveProjectFromSheetMetadata()
        {
            var sheetName = activeSheetNameProvider.Invoke() ?? string.Empty;
            RefreshProjectFromSheetMetadata(sheetName);
        }

        internal void RefreshProjectFromSheetMetadata(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                lastRefreshedSheetName = string.Empty;
                ClearActiveProjectState();
                return;
            }

            if (string.Equals(lastRefreshedSheetName, sheetName, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            try
            {
                var binding = metadataStore.LoadBinding(sheetName);
                lastRefreshedSheetName = sheetName;
                ApplyBindingState(binding);
            }
            catch (InvalidOperationException)
            {
                lastRefreshedSheetName = sheetName;
                ClearActiveProjectState();
            }
        }

        internal void InvalidateRefreshState()
        {
            lastRefreshedSheetName = string.Empty;
        }

        public void ExecuteFullDownload()
        {
            ExecuteDownload(service => service.PrepareFullDownload(GetRequiredSheetName()));
        }

        public void ExecutePartialDownload()
        {
            ExecuteDownload(service => service.PreparePartialDownload(GetRequiredSheetName()));
        }

        public void ExecuteFullUpload()
        {
            ExecuteUpload(service => service.PrepareFullUpload(GetRequiredSheetName()));
        }

        public void ExecutePartialUpload()
        {
            ExecuteUpload(service => service.PreparePartialUpload(GetRequiredSheetName()));
        }

        public void ExecuteInitializeCurrentSheet()
        {
            if (!EnsureProjectSelected())
            {
                return;
            }

            var sheetName = string.Empty;
            var project = new ProjectOption
            {
                SystemKey = ActiveSystemKey,
                ProjectId = ActiveProjectId,
                DisplayName = ActiveProjectDisplayName,
            };
            try
            {
                sheetName = GetRequiredSheetName();
                OfficeAgentLog.Info(
                    "ribbon_sync",
                    "initialize_sheet.begin",
                    "Initializing current worksheet.",
                    BuildInitializeSheetDetails(sheetName, project));
                EnsureExecutionService().InitializeCurrentSheet(sheetName, project);
                OfficeAgentLog.Info(
                    "ribbon_sync",
                    "initialize_sheet.completed",
                    "Current worksheet initialized.",
                    BuildInitializeSheetDetails(sheetName, project));
                TrackRibbonEvent("ribbon.initialize.completed");
                dialogService.ShowInfo(GetStrings().InitializeCurrentSheetCompletedMessage);
            }
            catch (AuthenticationRequiredException ex)
            {
                TrackRibbonEvent("ribbon.initialize.failed", error: CreateOperationFailedError(ex));
                OfficeAgentLog.Warn(
                    "ribbon_sync",
                    "initialize_sheet.authentication_required",
                    "Authentication is required while initializing current worksheet.",
                    BuildInitializeSheetDetails(sheetName, project));
                HandleAuthenticationRequired(ex);
            }
            catch (Exception ex)
            {
                TrackRibbonEvent("ribbon.initialize.failed", error: CreateOperationFailedError(ex));
                OfficeAgentLog.Error(
                    "ribbon_sync",
                    "initialize_sheet.failed",
                    "Failed to initialize current worksheet.",
                    ex,
                    BuildInitializeSheetDetails(sheetName, project));
                dialogService.ShowError(ex.Message);
            }
        }

        public void ExecuteAiColumnMapping()
        {
            if (!EnsureProjectSelected())
            {
                return;
            }

            try
            {
                var strings = GetStrings();
                var sheetName = GetRequiredSheetName();
                var service = EnsureExecutionService();
                var preview = dialogService.RunAiColumnMappingWithProgress(
                    cancellationToken => service.PrepareAiColumnMappingPreviewAsync(sheetName, cancellationToken));
                if (preview == null)
                {
                    return;
                }

                if (!HasApplicableAiColumnMappings(preview))
                {
                    TrackRibbonEvent(
                        "ribbon.ai_map_columns.completed",
                        new Dictionary<string, object>(StringComparer.Ordinal)
                        {
                            ["appliedCount"] = 0,
                            ["skippedCount"] = preview?.Items?.Length ?? 0,
                        });
                    dialogService.ShowInfo(strings.AiColumnMappingNoAcceptedMappingsMessage);
                    return;
                }

                if (!dialogService.ConfirmAiColumnMapping(preview))
                {
                    return;
                }

                var result = service.ApplyAiColumnMappingPreview(sheetName, preview);
                TrackRibbonEvent(
                    "ribbon.ai_map_columns.completed",
                    new Dictionary<string, object>(StringComparer.Ordinal)
                    {
                        ["appliedCount"] = result.AppliedCount,
                        ["skippedCount"] = result.SkippedCount,
                    });
                dialogService.ShowInfo(result.AppliedCount == 0
                    ? strings.AiColumnMappingNoAcceptedMappingsMessage
                    : strings.AiColumnMappingCompletedMessage(result.AppliedCount, result.SkippedCount));
            }
            catch (AuthenticationRequiredException ex)
            {
                TrackRibbonEvent("ribbon.ai_map_columns.failed", error: CreateOperationFailedError(ex));
                HandleAuthenticationRequired(ex);
            }
            catch (Exception ex)
            {
                TrackRibbonEvent("ribbon.ai_map_columns.failed", error: CreateOperationFailedError(ex));
                dialogService.ShowError(ex.Message);
            }
        }

        private static bool HasApplicableAiColumnMappings(AiColumnMappingPreview preview)
        {
            return (preview?.Items ?? Array.Empty<AiColumnMappingPreviewItem>())
                .Any(item => item != null &&
                             string.Equals(item.Status, AiColumnMappingPreviewStatuses.Accepted, StringComparison.Ordinal));
        }

        private void ExecuteDownload(Func<WorksheetSyncExecutionService, WorksheetDownloadPlan> preparePlan)
        {
            if (!EnsureProjectSelected())
            {
                return;
            }

            try
            {
                var strings = GetStrings();
                var plan = preparePlan(EnsureExecutionService());
                var rowCount = plan.Rows?.Count ?? 0;
                if (rowCount == 0)
                {
                    TrackRibbonEvent(
                        "ribbon.download.completed",
                        BuildDownloadProperties(plan, rowCount, fieldCount: 0));
                    dialogService.ShowInfo(strings.FormatDownloadNoMatchingRowsMessage(plan.OperationName));
                    return;
                }

                var fieldCount = CountDownloadFields(plan);
                if (!dialogService.ConfirmDownload(
                        strings.LocalizeSyncOperationName(plan.OperationName),
                        ActiveProjectDisplayName,
                        rowCount,
                        fieldCount,
                        plan.Preview))
                {
                    TrackRibbonEvent("ribbon.download.canceled", BuildDownloadProperties(plan, rowCount, fieldCount));
                    return;
                }

                TrackRibbonEvent("ribbon.download.confirmed", BuildDownloadProperties(plan, rowCount, fieldCount));
                executionService.ExecuteDownload(plan);
                TrackRibbonEvent("ribbon.download.completed", BuildDownloadProperties(plan, rowCount, fieldCount));
                dialogService.ShowInfo(strings.FormatDownloadCompletedMessage(
                    plan.OperationName,
                    rowCount,
                    fieldCount));
            }
            catch (AuthenticationRequiredException ex)
            {
                TrackRibbonEvent("ribbon.download.failed", error: CreateOperationFailedError(ex));
                HandleAuthenticationRequired(ex);
            }
            catch (Exception ex)
            {
                TrackRibbonEvent("ribbon.download.failed", error: CreateOperationFailedError(ex));
                dialogService.ShowError(ex.Message);
            }
        }

        private void ExecuteUpload(Func<WorksheetSyncExecutionService, WorksheetUploadPlan> preparePlan)
        {
            if (!EnsureProjectSelected())
            {
                return;
            }

            try
            {
                var strings = GetStrings();
                var plan = preparePlan(EnsureExecutionService());
                var preview = plan.Preview ?? new SyncOperationPreview();
                TrackRibbonEvent("ribbon.upload.previewed", BuildUploadProperties(plan, preview));
                if (preview.Changes.Length == 0)
                {
                    TrackRibbonEvent("ribbon.upload.completed", BuildUploadProperties(plan, preview));
                    dialogService.ShowInfo(preview.SkippedChanges.Length == 0
                        ? strings.FormatUploadNoChangesMessage(plan.OperationName)
                        : BuildUploadPreviewInfoMessage(strings, plan.OperationName, preview));
                    return;
                }

                if (!dialogService.ConfirmUpload(strings.LocalizeSyncOperationName(plan.OperationName), ActiveProjectDisplayName, preview))
                {
                    TrackRibbonEvent("ribbon.upload.canceled", BuildUploadProperties(plan, preview));
                    return;
                }

                TrackRibbonEvent("ribbon.upload.confirmed", BuildUploadProperties(plan, preview));
                executionService.ExecuteUpload(plan);
                TrackRibbonEvent("ribbon.upload.completed", BuildUploadProperties(plan, preview));
                dialogService.ShowInfo(BuildUploadCompletionMessage(strings, plan.OperationName, preview));
            }
            catch (AuthenticationRequiredException ex)
            {
                TrackRibbonEvent("ribbon.upload.failed", error: CreateOperationFailedError(ex));
                HandleAuthenticationRequired(ex);
            }
            catch (Exception ex)
            {
                TrackRibbonEvent("ribbon.upload.failed", error: CreateOperationFailedError(ex));
                dialogService.ShowError(ex.Message);
            }
        }

        private static string BuildUploadPreviewInfoMessage(
            HostLocalizedStrings strings,
            string operationName,
            SyncOperationPreview preview)
        {
            var builder = new StringBuilder();
            var summary = preview?.Summary ?? string.Empty;
            builder.AppendLine(string.IsNullOrWhiteSpace(summary)
                ? strings.FormatUploadNoChangesMessage(operationName)
                : summary);

            foreach (var detail in preview?.Details ?? Array.Empty<string>())
            {
                if (!string.IsNullOrWhiteSpace(detail))
                {
                    builder.AppendLine(detail);
                }
            }

            return builder.ToString().TrimEnd();
        }

        private static string BuildUploadCompletionMessage(
            HostLocalizedStrings strings,
            string operationName,
            SyncOperationPreview preview)
        {
            var builder = new StringBuilder(strings.FormatUploadCompletedMessage(
                operationName,
                preview?.Changes?.Length ?? 0));

            var skippedCount = preview?.SkippedChanges?.Length ?? 0;
            if (skippedCount > 0)
            {
                builder
                    .AppendLine()
                    .Append(strings.SkippedCellCountLine(skippedCount));
            }

            return builder.ToString();
        }

        private bool EnsureProjectSelected()
        {
            if (!string.IsNullOrWhiteSpace(ActiveProjectId))
            {
                return true;
            }

            dialogService.ShowWarning(GetStrings().ProjectSelectionRequiredMessage);
            return false;
        }

        private void ApplyBindingState(SheetBinding binding)
        {
            ActiveBinding = binding;
            ActiveProjectId = binding?.ProjectId ?? string.Empty;
            ActiveSystemKey = binding?.SystemKey ?? string.Empty;
            ActiveProjectDisplayName = string.IsNullOrWhiteSpace(binding?.ProjectName)
                ? string.Empty
                : binding.ProjectName;
            OnActiveProjectChanged();
        }

        private void ClearActiveProjectState()
        {
            ActiveBinding = null;
            ActiveProjectId = string.Empty;
            ActiveSystemKey = string.Empty;
            ActiveProjectDisplayName = GetStrings().ProjectDropDownPlaceholderText;
            OnActiveProjectChanged();
        }

        private static HostLocalizedStrings GetStrings()
        {
            return Globals.ThisAddIn?.HostLocalizedStrings ?? HostLocalizedStrings.ForLocale("en");
        }

        private void OnActiveProjectChanged()
        {
            ActiveProjectChanged?.Invoke(this, EventArgs.Empty);
        }

        private WorksheetSyncExecutionService EnsureExecutionService()
        {
            if (executionService == null)
            {
                throw new InvalidOperationException("Worksheet sync execution service is not configured.");
            }

            return executionService;
        }

        private string GetRequiredSheetName()
        {
            var sheetName = activeSheetNameProvider.Invoke() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                throw new InvalidOperationException("Active worksheet is not available.");
            }

            return sheetName;
        }

        private SheetBinding TryLoadBinding(string sheetName)
        {
            try
            {
                return metadataStore.LoadBinding(sheetName);
            }
            catch (InvalidOperationException)
            {
                return null;
            }
        }

        private static bool IsSameProject(SheetBinding existingBinding, ProjectOption project)
        {
            if (existingBinding == null || project == null)
            {
                return false;
            }

            return string.Equals(existingBinding.SystemKey, project.SystemKey, StringComparison.Ordinal) &&
                   string.Equals(existingBinding.ProjectId, project.ProjectId, StringComparison.Ordinal);
        }

        private void RestoreBindingState(SheetBinding binding, string sheetName)
        {
            lastRefreshedSheetName = sheetName;
            if (binding == null)
            {
                ClearActiveProjectState();
                return;
            }

            ApplyBindingState(binding);
        }

        private void HandleAuthenticationRequired(AuthenticationRequiredException ex)
        {
            if (dialogService.ShowAuthenticationRequired(GetStrings().AuthenticationRequiredDefaultMessage))
            {
                authenticationLoginAction?.Invoke();
            }
        }

        private void TrackRibbonEvent(
            string eventName,
            IDictionary<string, object> properties = null,
            AnalyticsError error = null)
        {
            if (string.IsNullOrWhiteSpace(eventName))
            {
                return;
            }

            var strings = GetStrings();
            var sheetName = SafeGetActiveSheetName();
            var merged = new Dictionary<string, object>(StringComparer.Ordinal)
            {
                ["systemKey"] = ActiveSystemKey ?? string.Empty,
                ["projectId"] = ActiveProjectId ?? string.Empty,
                ["projectName"] = ActiveProjectDisplayName ?? string.Empty,
                ["sheetName"] = sheetName,
                ["uiLocale"] = strings.Locale ?? string.Empty,
            };

            if (properties != null)
            {
                foreach (var property in properties)
                {
                    merged[property.Key ?? string.Empty] = property.Value;
                }
            }

            analyticsService.Track(eventName, "ribbon", merged, error: error);
        }

        private string SafeGetActiveSheetName()
        {
            try
            {
                return activeSheetNameProvider.Invoke() ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static AnalyticsError CreateOperationFailedError(Exception ex)
        {
            return new AnalyticsError
            {
                Code = "operation_failed",
                Message = ex?.Message ?? string.Empty,
                ExceptionType = ex?.GetType().Name ?? string.Empty,
            };
        }

        private static IDictionary<string, object> BuildDownloadProperties(
            WorksheetDownloadPlan plan,
            int rowCount,
            int fieldCount)
        {
            return new Dictionary<string, object>(StringComparer.Ordinal)
            {
                ["operationName"] = plan?.OperationName ?? string.Empty,
                ["operationScope"] = GetOperationScope(plan?.OperationName),
                ["rowCount"] = rowCount,
                ["fieldCount"] = fieldCount,
            };
        }

        private static IDictionary<string, object> BuildUploadProperties(
            WorksheetUploadPlan plan,
            SyncOperationPreview preview)
        {
            return new Dictionary<string, object>(StringComparer.Ordinal)
            {
                ["operationName"] = plan?.OperationName ?? string.Empty,
                ["operationScope"] = GetOperationScope(plan?.OperationName),
                ["submittedCellCount"] = preview?.Changes?.Length ?? 0,
                ["skippedCellCount"] = preview?.SkippedChanges?.Length ?? 0,
            };
        }

        private static string GetOperationScope(string operationName)
        {
            if (string.IsNullOrWhiteSpace(operationName))
            {
                return string.Empty;
            }

            return operationName.IndexOf("全量", StringComparison.Ordinal) >= 0
                ? "full"
                : "partial";
        }

        private static int CountDownloadFields(WorksheetDownloadPlan plan)
        {
            if (plan?.Selection?.TargetColumns?.Length > 0)
            {
                return plan.Selection.TargetColumns
                    .Distinct()
                    .Count();
            }

            if (plan?.Selection?.TargetCells?.Length > 0)
            {
                return plan.Selection.TargetCells
                    .Select(cell => cell.Column)
                    .Distinct()
                    .Count();
            }

            return plan?.Schema?.Columns?.Length ?? 0;
        }

        private static string BuildProjectSelectionDetails(
            string sheetName,
            ProjectOption targetProject,
            SheetBinding existingBinding,
            SheetBinding layoutBinding = null,
            bool includeRestoredProject = false)
        {
            var builder = new StringBuilder();
            AppendDetail(builder, "SheetName", sheetName);
            AppendDetail(builder, "TargetSystemKey", targetProject?.SystemKey);
            AppendDetail(builder, "TargetProjectId", targetProject?.ProjectId);
            AppendDetail(builder, "TargetProjectName", targetProject?.DisplayName);
            AppendDetail(builder, "ExistingSystemKey", existingBinding?.SystemKey);
            AppendDetail(builder, "ExistingProjectId", existingBinding?.ProjectId);
            AppendDetail(builder, "ExistingProjectName", existingBinding?.ProjectName);

            if (layoutBinding != null)
            {
                if (includeRestoredProject)
                {
                    AppendDetail(builder, "RestoredProjectId", existingBinding?.ProjectId);
                }

                AppendDetail(builder, "HeaderStartRow", layoutBinding.HeaderStartRow.ToString());
                AppendDetail(builder, "HeaderRowCount", layoutBinding.HeaderRowCount.ToString());
                AppendDetail(builder, "DataStartRow", layoutBinding.DataStartRow.ToString());
            }

            return builder.ToString();
        }

        private static string BuildInitializeSheetDetails(string sheetName, ProjectOption project)
        {
            var builder = new StringBuilder();
            AppendDetail(builder, "SheetName", sheetName);
            AppendDetail(builder, "SystemKey", project?.SystemKey);
            AppendDetail(builder, "ProjectId", project?.ProjectId);
            AppendDetail(builder, "ProjectName", project?.DisplayName);
            return builder.ToString();
        }

        private static void AppendDetail(StringBuilder builder, string name, string value)
        {
            if (builder.Length > 0)
            {
                builder.Append("; ");
            }

            builder
                .Append(name)
                .Append('=')
                .Append(string.IsNullOrWhiteSpace(value) ? "<empty>" : value);
        }
    }
}
