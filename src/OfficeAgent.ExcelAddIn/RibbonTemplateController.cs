using System;
using System.Linq;
using System.Windows.Forms;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
using OfficeAgent.ExcelAddIn.Dialogs;

namespace OfficeAgent.ExcelAddIn
{
    internal sealed class RibbonTemplateController
    {
        private const string DefaultTemplateDisplayName = "未绑定模板";

        private readonly ITemplateCatalog templateCatalog;
        private readonly Func<string> activeSheetNameProvider;
        private readonly IRibbonTemplateDialogService dialogService;
        private string lastRefreshedSheetName = string.Empty;

        public RibbonTemplateController(
            ITemplateCatalog templateCatalog,
            Func<string> activeSheetNameProvider)
            : this(templateCatalog, activeSheetNameProvider, new RibbonTemplateDialogService())
        {
        }

        internal RibbonTemplateController(
            ITemplateCatalog templateCatalog,
            Func<string> activeSheetNameProvider,
            IRibbonTemplateDialogService dialogService)
        {
            this.templateCatalog = templateCatalog ?? throw new ArgumentNullException(nameof(templateCatalog));
            this.activeSheetNameProvider = activeSheetNameProvider ?? throw new ArgumentNullException(nameof(activeSheetNameProvider));
            this.dialogService = dialogService ?? throw new ArgumentNullException(nameof(dialogService));

            ActiveTemplateDisplayName = DefaultTemplateDisplayName;
        }

        public event EventHandler TemplateStateChanged;

        public string ActiveTemplateDisplayName { get; private set; }

        public bool CanApplyTemplate { get; private set; }

        public bool CanSaveTemplate { get; private set; }

        public bool CanSaveAsTemplate { get; private set; }

        public void RefreshActiveTemplateStateFromSheetMetadata()
        {
            RefreshTemplateState(activeSheetNameProvider.Invoke() ?? string.Empty);
        }

        internal void RefreshTemplateState(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                lastRefreshedSheetName = string.Empty;
                ApplyState(new SheetTemplateState());
                return;
            }

            if (string.Equals(lastRefreshedSheetName, sheetName, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            lastRefreshedSheetName = sheetName;
            ApplyState(templateCatalog.GetSheetState(sheetName) ?? new SheetTemplateState());
        }

        internal void InvalidateRefreshState()
        {
            lastRefreshedSheetName = string.Empty;
        }

        public void ExecuteApplyTemplate()
        {
            try
            {
                var sheetName = GetRequiredSheetName();
                var state = templateCatalog.GetSheetState(sheetName) ?? new SheetTemplateState();
                if (!state.CanApplyTemplate)
                {
                    dialogService.ShowWarning("请先选择项目。");
                    return;
                }

                var templates = templateCatalog.ListTemplates(sheetName) ?? Array.Empty<TemplateDefinition>();
                if (templates.Count == 0)
                {
                    dialogService.ShowWarning("当前项目没有可用模板。");
                    return;
                }

                var templateId = dialogService.ShowTemplatePicker(state.ProjectDisplayName, templates);
                if (string.IsNullOrWhiteSpace(templateId))
                {
                    return;
                }

                var selectedTemplate = templates.FirstOrDefault(template =>
                    string.Equals(template.TemplateId, templateId, StringComparison.Ordinal));
                if (selectedTemplate == null)
                {
                    dialogService.ShowWarning("未找到所选模板。");
                    return;
                }

                if (state.IsDirty && !dialogService.ConfirmApplyTemplateOverwrite(selectedTemplate.TemplateName))
                {
                    return;
                }

                templateCatalog.ApplyTemplateToSheet(sheetName, templateId);
                InvalidateRefreshState();
                RefreshActiveTemplateStateFromSheetMetadata();
                dialogService.ShowInfo($"应用模板完成。\r\n模板：{selectedTemplate.TemplateName}");
            }
            catch (Exception ex)
            {
                dialogService.ShowError(ex.Message);
            }
        }

        public void ExecuteSaveTemplate()
        {
            try
            {
                var sheetName = GetRequiredSheetName();
                var state = templateCatalog.GetSheetState(sheetName) ?? new SheetTemplateState();
                if (!state.CanSaveTemplate || string.IsNullOrWhiteSpace(state.TemplateId) || !state.TemplateRevision.HasValue)
                {
                    dialogService.ShowWarning("当前表没有可保存的模板。");
                    return;
                }

                if (TrySaveTemplate(sheetName, state, overwriteRevisionConflict: false, $"保存模板完成。\r\n模板：{state.TemplateName}"))
                {
                    return;
                }

                var conflictResult = dialogService.ShowTemplateRevisionConflictDialog(
                    state.TemplateName,
                    state.TemplateRevision.Value,
                    state.StoredTemplateRevision ?? state.TemplateRevision.Value);

                if (conflictResult == DialogResult.Yes)
                {
                    TrySaveTemplate(sheetName, state, overwriteRevisionConflict: true, $"覆盖模板完成。\r\n模板：{state.TemplateName}");
                    return;
                }

                if (conflictResult == DialogResult.No)
                {
                    ExecuteSaveAsTemplate();
                }
            }
            catch (Exception ex)
            {
                dialogService.ShowError(ex.Message);
            }
        }

        public void ExecuteSaveAsTemplate()
        {
            try
            {
                var sheetName = GetRequiredSheetName();
                var state = templateCatalog.GetSheetState(sheetName) ?? new SheetTemplateState();
                if (!state.CanSaveAsTemplate)
                {
                    dialogService.ShowWarning("请先选择项目。");
                    return;
                }

                var suggestedTemplateName = string.IsNullOrWhiteSpace(state.TemplateName)
                    ? "新模板"
                    : state.TemplateName + "-副本";
                var templateName = dialogService.ShowSaveAsTemplateDialog(suggestedTemplateName);
                if (string.IsNullOrWhiteSpace(templateName))
                {
                    return;
                }

                templateCatalog.SaveSheetAsNewTemplate(sheetName, templateName);
                InvalidateRefreshState();
                RefreshActiveTemplateStateFromSheetMetadata();
                dialogService.ShowInfo($"另存模板完成。\r\n模板：{templateName}");
            }
            catch (Exception ex)
            {
                dialogService.ShowError(ex.Message);
            }
        }

        private bool TrySaveTemplate(
            string sheetName,
            SheetTemplateState state,
            bool overwriteRevisionConflict,
            string successMessage)
        {
            try
            {
                templateCatalog.SaveSheetToExistingTemplate(
                    sheetName,
                    state.TemplateId,
                    state.TemplateRevision ?? 0,
                    overwriteRevisionConflict);
                InvalidateRefreshState();
                RefreshActiveTemplateStateFromSheetMetadata();
                dialogService.ShowInfo(successMessage);
                return true;
            }
            catch (InvalidOperationException ex) when (!overwriteRevisionConflict && IsRevisionConflict(ex.Message))
            {
                return false;
            }
        }

        private void ApplyState(SheetTemplateState state)
        {
            CanApplyTemplate = state?.CanApplyTemplate == true;
            CanSaveTemplate = state?.CanSaveTemplate == true;
            CanSaveAsTemplate = state?.CanSaveAsTemplate == true;
            ActiveTemplateDisplayName = string.IsNullOrWhiteSpace(state?.TemplateName)
                ? DefaultTemplateDisplayName
                : state.TemplateName;
            TemplateStateChanged?.Invoke(this, EventArgs.Empty);
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

        private static bool IsRevisionConflict(string message)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                return false;
            }

            return string.Equals(message, "模板版本已变化。", StringComparison.Ordinal) ||
                   message.IndexOf("revision", StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }
}
