using System;
using System.Collections.Generic;
using System.Linq;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
using OfficeAgent.Core.Sync;
using OfficeAgent.ExcelAddIn.Dialogs;

namespace OfficeAgent.ExcelAddIn
{
    internal sealed class RibbonSyncController
    {
        private const string DefaultProjectDisplayName = "先选择项目";

        private readonly ISystemConnector connector;
        private readonly IWorksheetMetadataStore metadataStore;
        private readonly WorksheetSyncService worksheetSyncService;
        private readonly Func<string> activeSheetNameProvider;
        private readonly WorksheetSyncExecutionService executionService;

        public RibbonSyncController(
            ISystemConnector connector,
            IWorksheetMetadataStore metadataStore,
            WorksheetSyncService worksheetSyncService,
            Func<string> activeSheetNameProvider)
            : this(connector, metadataStore, worksheetSyncService, activeSheetNameProvider, executionService: null)
        {
        }

        internal RibbonSyncController(
            ISystemConnector connector,
            IWorksheetMetadataStore metadataStore,
            WorksheetSyncService worksheetSyncService,
            Func<string> activeSheetNameProvider,
            WorksheetSyncExecutionService executionService)
        {
            this.connector = connector ?? throw new ArgumentNullException(nameof(connector));
            this.metadataStore = metadataStore ?? throw new ArgumentNullException(nameof(metadataStore));
            this.worksheetSyncService = worksheetSyncService ?? throw new ArgumentNullException(nameof(worksheetSyncService));
            this.activeSheetNameProvider = activeSheetNameProvider ?? throw new ArgumentNullException(nameof(activeSheetNameProvider));
            this.executionService = executionService;

            ActiveProjectDisplayName = DefaultProjectDisplayName;
            ActiveProjectId = string.Empty;
            ActiveSystemKey = string.Empty;
        }

        public event EventHandler ActiveProjectChanged;

        public string ActiveProjectDisplayName { get; private set; }

        public string ActiveProjectId { get; private set; }

        public string ActiveSystemKey { get; private set; }

        public IReadOnlyList<ProjectOption> GetProjects()
        {
            return connector.GetProjects() ?? Array.Empty<ProjectOption>();
        }

        public void SelectProject(ProjectOption project)
        {
            if (project == null)
            {
                throw new ArgumentNullException(nameof(project));
            }

            var sheetName = activeSheetNameProvider.Invoke() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                throw new InvalidOperationException("Active worksheet is not available.");
            }

            var binding = new SheetBinding
            {
                SheetName = sheetName,
                SystemKey = project.SystemKey ?? string.Empty,
                ProjectId = project.ProjectId ?? string.Empty,
                ProjectName = project.DisplayName ?? string.Empty,
            };

            metadataStore.SaveBinding(binding);
            ApplyBindingState(binding);
        }

        public void RefreshActiveProjectFromSheetMetadata()
        {
            var sheetName = activeSheetNameProvider.Invoke() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                ClearActiveProjectState();
                return;
            }

            try
            {
                var binding = metadataStore.LoadBinding(sheetName);
                ApplyBindingState(binding);
            }
            catch (InvalidOperationException)
            {
                ClearActiveProjectState();
            }
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

        public void ExecuteIncrementalUpload()
        {
            ExecuteUpload(service => service.PrepareIncrementalUpload(GetRequiredSheetName()));
        }

        private void ExecuteDownload(Func<WorksheetSyncExecutionService, WorksheetDownloadPlan> preparePlan)
        {
            if (!EnsureProjectSelected())
            {
                return;
            }

            try
            {
                var plan = preparePlan(EnsureExecutionService());
                if (!DownloadConfirmDialog.Confirm(
                        plan.OperationName,
                        ActiveProjectDisplayName,
                        plan.Rows?.Count ?? 0,
                        CountDownloadFields(plan),
                        plan.Preview))
                {
                    return;
                }

                executionService.ExecuteDownload(plan);
                OperationResultDialog.ShowInfo(
                    $"{plan.OperationName}完成。\r\n记录数：{plan.Rows?.Count ?? 0}\r\n字段数：{CountDownloadFields(plan)}");
            }
            catch (Exception ex)
            {
                OperationResultDialog.ShowError(ex.Message);
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
                var plan = preparePlan(EnsureExecutionService());
                var preview = plan.Preview ?? new SyncOperationPreview();
                if (preview.Changes.Length == 0)
                {
                    OperationResultDialog.ShowInfo($"{plan.OperationName}没有可提交的单元格。");
                    return;
                }

                if (!UploadConfirmDialog.Confirm(plan.OperationName, ActiveProjectDisplayName, preview))
                {
                    return;
                }

                executionService.ExecuteUpload(plan);
                OperationResultDialog.ShowInfo($"{plan.OperationName}完成。\r\n提交单元格数：{preview.Changes.Length}");
            }
            catch (Exception ex)
            {
                OperationResultDialog.ShowError(ex.Message);
            }
        }

        private bool EnsureProjectSelected()
        {
            if (!string.IsNullOrWhiteSpace(ActiveProjectId))
            {
                return true;
            }

            OperationResultDialog.ShowWarning("请先选择项目。");
            return false;
        }

        private void ApplyBindingState(SheetBinding binding)
        {
            ActiveProjectId = binding?.ProjectId ?? string.Empty;
            ActiveSystemKey = binding?.SystemKey ?? string.Empty;
            ActiveProjectDisplayName = string.IsNullOrWhiteSpace(binding?.ProjectName)
                ? DefaultProjectDisplayName
                : binding.ProjectName;
            OnActiveProjectChanged();
        }

        private void ClearActiveProjectState()
        {
            ActiveProjectId = string.Empty;
            ActiveSystemKey = string.Empty;
            ActiveProjectDisplayName = DefaultProjectDisplayName;
            OnActiveProjectChanged();
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
    }
}
