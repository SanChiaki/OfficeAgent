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
        private readonly IRibbonSyncDialogService dialogService;

        public RibbonSyncController(
            ISystemConnector connector,
            IWorksheetMetadataStore metadataStore,
            WorksheetSyncService worksheetSyncService,
            Func<string> activeSheetNameProvider)
            : this(
                connector,
                metadataStore,
                worksheetSyncService,
                activeSheetNameProvider,
                executionService: null,
                new RibbonSyncDialogService())
        {
        }

        internal RibbonSyncController(
            ISystemConnector connector,
            IWorksheetMetadataStore metadataStore,
            WorksheetSyncService worksheetSyncService,
            Func<string> activeSheetNameProvider,
            WorksheetSyncExecutionService executionService)
            : this(
                connector,
                metadataStore,
                worksheetSyncService,
                activeSheetNameProvider,
                executionService,
                new RibbonSyncDialogService())
        {
        }

        internal RibbonSyncController(
            ISystemConnector connector,
            IWorksheetMetadataStore metadataStore,
            WorksheetSyncService worksheetSyncService,
            Func<string> activeSheetNameProvider,
            WorksheetSyncExecutionService executionService,
            IRibbonSyncDialogService dialogService)
        {
            this.connector = connector ?? throw new ArgumentNullException(nameof(connector));
            this.metadataStore = metadataStore ?? throw new ArgumentNullException(nameof(metadataStore));
            this.worksheetSyncService = worksheetSyncService ?? throw new ArgumentNullException(nameof(worksheetSyncService));
            this.activeSheetNameProvider = activeSheetNameProvider ?? throw new ArgumentNullException(nameof(activeSheetNameProvider));
            this.executionService = executionService;
            this.dialogService = dialogService ?? throw new ArgumentNullException(nameof(dialogService));

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

            var binding = BuildBindingForSelection(sheetName, project);

            metadataStore.SaveBinding(binding);
            ApplyBindingState(binding);
            TryAutoInitializeCurrentSheet(sheetName, project);
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
                dialogService.ShowInfo("初始化当前表完成。");
            }
            catch (Exception ex)
            {
                dialogService.ShowError(ex.Message);
            }
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
                if (!dialogService.ConfirmDownload(
                        plan.OperationName,
                        ActiveProjectDisplayName,
                        plan.Rows?.Count ?? 0,
                        CountDownloadFields(plan),
                        plan.Preview))
                {
                    return;
                }

                executionService.ExecuteDownload(plan);
                dialogService.ShowInfo(
                    $"{plan.OperationName}完成。\r\n记录数：{plan.Rows?.Count ?? 0}\r\n字段数：{CountDownloadFields(plan)}");
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
                var plan = preparePlan(EnsureExecutionService());
                var preview = plan.Preview ?? new SyncOperationPreview();
                if (preview.Changes.Length == 0)
                {
                    dialogService.ShowInfo($"{plan.OperationName}没有可提交的单元格。");
                    return;
                }

                if (!dialogService.ConfirmUpload(plan.OperationName, ActiveProjectDisplayName, preview))
                {
                    return;
                }

                executionService.ExecuteUpload(plan);
                dialogService.ShowInfo($"{plan.OperationName}完成。\r\n提交单元格数：{preview.Changes.Length}");
            }
            catch (Exception ex)
            {
                dialogService.ShowError(ex.Message);
            }
        }

        private bool EnsureProjectSelected()
        {
            if (!string.IsNullOrWhiteSpace(ActiveProjectId))
            {
                return true;
            }

            dialogService.ShowWarning("请先选择项目。");
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

        private SheetBinding BuildBindingForSelection(string sheetName, ProjectOption project)
        {
            var seed = connector.CreateBindingSeed(sheetName, project);

            try
            {
                var existing = metadataStore.LoadBinding(sheetName);
                return new SheetBinding
                {
                    SheetName = sheetName,
                    SystemKey = seed.SystemKey,
                    ProjectId = seed.ProjectId,
                    ProjectName = seed.ProjectName,
                    HeaderStartRow = existing.HeaderStartRow > 0 ? existing.HeaderStartRow : seed.HeaderStartRow,
                    HeaderRowCount = existing.HeaderRowCount > 0 ? existing.HeaderRowCount : seed.HeaderRowCount,
                    DataStartRow = existing.DataStartRow > 0 ? existing.DataStartRow : seed.DataStartRow,
                };
            }
            catch (InvalidOperationException)
            {
                return seed;
            }
        }

        private void TryAutoInitializeCurrentSheet(string sheetName, ProjectOption project)
        {
            if (executionService == null)
            {
                return;
            }

            try
            {
                executionService.TryAutoInitializeCurrentSheet(sheetName, project);
            }
            catch (Exception ex)
            {
                dialogService.ShowWarning($"自动初始化未完成，请使用“初始化当前表”并检查 _OfficeAgentMetadata。\r\n{ex.Message}");
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
    }
}
