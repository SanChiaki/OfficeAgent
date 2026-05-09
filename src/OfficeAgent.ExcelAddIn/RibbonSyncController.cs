using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeAgent.Core;
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
        {
            this.metadataStore = metadataStore ?? throw new ArgumentNullException(nameof(metadataStore));
            this.worksheetSyncService = worksheetSyncService ?? throw new ArgumentNullException(nameof(worksheetSyncService));
            this.activeSheetNameProvider = activeSheetNameProvider ?? throw new ArgumentNullException(nameof(activeSheetNameProvider));
            this.executionService = executionService;
            this.dialogService = dialogService ?? throw new ArgumentNullException(nameof(dialogService));
            this.authenticationLoginAction = authenticationLoginAction;

            ActiveProjectDisplayName = GetStrings().ProjectDropDownPlaceholderText;
            ActiveProjectId = string.Empty;
            ActiveSystemKey = string.Empty;
        }

        public event EventHandler ActiveProjectChanged;

        public string ActiveProjectDisplayName { get; private set; }

        public string ActiveProjectId { get; private set; }

        public string ActiveSystemKey { get; private set; }

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
                return;
            }

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

            try
            {
                var sheetName = GetRequiredSheetName();
                EnsureExecutionService().InitializeCurrentSheet(sheetName, new ProjectOption
                {
                    SystemKey = ActiveSystemKey,
                    ProjectId = ActiveProjectId,
                    DisplayName = ActiveProjectDisplayName,
                });
                dialogService.ShowInfo(GetStrings().InitializeCurrentSheetCompletedMessage);
            }
            catch (AuthenticationRequiredException ex)
            {
                HandleAuthenticationRequired(ex);
            }
            catch (Exception ex)
            {
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
                    dialogService.ShowInfo(strings.AiColumnMappingNoAcceptedMappingsMessage);
                    return;
                }

                if (!dialogService.ConfirmAiColumnMapping(preview))
                {
                    return;
                }

                var result = service.ApplyAiColumnMappingPreview(sheetName, preview);
                dialogService.ShowInfo(result.AppliedCount == 0
                    ? strings.AiColumnMappingNoAcceptedMappingsMessage
                    : strings.AiColumnMappingCompletedMessage(result.AppliedCount, result.SkippedCount));
            }
            catch (AuthenticationRequiredException ex)
            {
                HandleAuthenticationRequired(ex);
            }
            catch (Exception ex)
            {
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
                if (!dialogService.ConfirmDownload(
                        strings.LocalizeSyncOperationName(plan.OperationName),
                        ActiveProjectDisplayName,
                        plan.Rows?.Count ?? 0,
                        CountDownloadFields(plan),
                        plan.Preview))
                {
                    return;
                }

                executionService.ExecuteDownload(plan);
                dialogService.ShowInfo(strings.FormatDownloadCompletedMessage(
                    plan.OperationName,
                    plan.Rows?.Count ?? 0,
                    CountDownloadFields(plan)));
            }
            catch (AuthenticationRequiredException ex)
            {
                HandleAuthenticationRequired(ex);
            }
            catch (Exception ex)
            {
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
                if (preview.Changes.Length == 0)
                {
                    dialogService.ShowInfo(preview.SkippedChanges.Length == 0
                        ? strings.FormatUploadNoChangesMessage(plan.OperationName)
                        : BuildUploadPreviewInfoMessage(strings, plan.OperationName, preview));
                    return;
                }

                if (!dialogService.ConfirmUpload(strings.LocalizeSyncOperationName(plan.OperationName), ActiveProjectDisplayName, preview))
                {
                    return;
                }

                executionService.ExecuteUpload(plan);
                dialogService.ShowInfo(BuildUploadCompletionMessage(strings, plan.OperationName, preview));
            }
            catch (AuthenticationRequiredException ex)
            {
                HandleAuthenticationRequired(ex);
            }
            catch (Exception ex)
            {
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
            ActiveProjectId = binding?.ProjectId ?? string.Empty;
            ActiveSystemKey = binding?.SystemKey ?? string.Empty;
            ActiveProjectDisplayName = string.IsNullOrWhiteSpace(binding?.ProjectName)
                ? string.Empty
                : binding.ProjectName;
            OnActiveProjectChanged();
        }

        private void ClearActiveProjectState()
        {
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

        private static int CountDownloadFields(WorksheetDownloadPlan plan)
        {
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
