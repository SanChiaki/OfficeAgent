using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using OfficeAgent.Core.Analytics;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
using OfficeAgent.ExcelAddIn.Dialogs;
using OfficeAgent.ExcelAddIn.Localization;

namespace OfficeAgent.ExcelAddIn
{
    internal sealed class RibbonTemplateController
    {
        private readonly ITemplateCatalog templateCatalog;
        private readonly Func<string> activeSheetNameProvider;
        private readonly IRibbonTemplateDialogService dialogService;
        private readonly IAnalyticsService analyticsService;
        private string lastRefreshedSheetName = string.Empty;

        public RibbonTemplateController(
            ITemplateCatalog templateCatalog,
            Func<string> activeSheetNameProvider)
            : this(
                templateCatalog,
                activeSheetNameProvider,
                new RibbonTemplateDialogService(),
                NoopAnalyticsService.Instance)
        {
        }

        internal RibbonTemplateController(
            ITemplateCatalog templateCatalog,
            Func<string> activeSheetNameProvider,
            IRibbonTemplateDialogService dialogService)
            : this(templateCatalog, activeSheetNameProvider, dialogService, analyticsService: null)
        {
        }

        internal RibbonTemplateController(
            ITemplateCatalog templateCatalog,
            Func<string> activeSheetNameProvider,
            IRibbonTemplateDialogService dialogService,
            IAnalyticsService analyticsService = null)
        {
            this.templateCatalog = templateCatalog ?? throw new ArgumentNullException(nameof(templateCatalog));
            this.activeSheetNameProvider = activeSheetNameProvider ?? throw new ArgumentNullException(nameof(activeSheetNameProvider));
            this.dialogService = dialogService ?? throw new ArgumentNullException(nameof(dialogService));
            this.analyticsService = analyticsService ?? NoopAnalyticsService.Instance;

            ActiveTemplateDisplayName = GetStrings().DefaultTemplateDisplayName;
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
            SheetTemplateState state = null;
            TemplateDefinition selectedTemplate = null;
            try
            {
                var sheetName = GetRequiredSheetName();
                state = templateCatalog.GetSheetState(sheetName) ?? new SheetTemplateState();
                TrackTemplateEvent("ribbon.template.apply.clicked", sheetName, state);
                if (!state.CanApplyTemplate)
                {
                    dialogService.ShowWarning(GetStrings().ProjectSelectionRequiredMessage);
                    return;
                }

                var templates = templateCatalog.ListTemplates(sheetName) ?? Array.Empty<TemplateDefinition>();
                if (templates.Count == 0)
                {
                    dialogService.ShowWarning(GetStrings().TemplateNoAvailableMessage);
                    return;
                }

                var templateId = dialogService.ShowTemplatePicker(state.ProjectDisplayName, templates);
                if (string.IsNullOrWhiteSpace(templateId))
                {
                    return;
                }

                selectedTemplate = templates.FirstOrDefault(template =>
                    string.Equals(template.TemplateId, templateId, StringComparison.Ordinal));
                if (selectedTemplate == null)
                {
                    dialogService.ShowWarning(GetStrings().TemplateNotFoundMessage);
                    return;
                }

                if (state.IsDirty && !dialogService.ConfirmApplyTemplateOverwrite(selectedTemplate.TemplateName))
                {
                    return;
                }

                templateCatalog.ApplyTemplateToSheet(sheetName, templateId);
                InvalidateRefreshState();
                RefreshActiveTemplateStateFromSheetMetadata();
                TrackTemplateEvent("ribbon.template.apply.completed", sheetName, state, selectedTemplate);
                dialogService.ShowInfo(GetStrings().ApplyTemplateCompletedMessage(selectedTemplate.TemplateName));
            }
            catch (Exception ex)
            {
                TrackTemplateEvent("ribbon.template.apply.failed", GetOptionalSheetName(), state, selectedTemplate, CreateTemplateOperationFailedError(ex));
                dialogService.ShowError(ex.Message);
            }
        }

        public void ExecuteSaveTemplate()
        {
            SheetTemplateState state = null;
            try
            {
                var sheetName = GetRequiredSheetName();
                state = templateCatalog.GetSheetState(sheetName) ?? new SheetTemplateState();
                TrackTemplateEvent("ribbon.template.save.clicked", sheetName, state);
                if (!state.CanSaveTemplate || string.IsNullOrWhiteSpace(state.TemplateId) || !state.TemplateRevision.HasValue)
                {
                    dialogService.ShowWarning(GetStrings().TemplateNoSavableMessage);
                    return;
                }

                if (TrySaveTemplate(sheetName, state, overwriteRevisionConflict: false, GetStrings().SaveTemplateCompletedMessage(state.TemplateName)))
                {
                    return;
                }

                var conflictResult = dialogService.ShowTemplateRevisionConflictDialog(
                    state.TemplateName,
                    state.TemplateRevision.Value,
                    state.StoredTemplateRevision ?? state.TemplateRevision.Value);

                if (conflictResult == DialogResult.Yes)
                {
                    TrySaveTemplate(sheetName, state, overwriteRevisionConflict: true, GetStrings().OverwriteTemplateCompletedMessage(state.TemplateName));
                    return;
                }

                if (conflictResult == DialogResult.No)
                {
                    ExecuteSaveAsTemplate();
                }
            }
            catch (Exception ex)
            {
                TrackTemplateEvent("ribbon.template.save.failed", GetOptionalSheetName(), state, error: CreateTemplateOperationFailedError(ex));
                dialogService.ShowError(ex.Message);
            }
        }

        public void ExecuteSaveAsTemplate()
        {
            SheetTemplateState state = null;
            string templateName = string.Empty;
            try
            {
                var sheetName = GetRequiredSheetName();
                state = templateCatalog.GetSheetState(sheetName) ?? new SheetTemplateState();
                TrackTemplateEvent("ribbon.template.save_as.clicked", sheetName, state);
                if (!state.CanSaveAsTemplate)
                {
                    dialogService.ShowWarning(GetStrings().ProjectSelectionRequiredMessage);
                    return;
                }

                var suggestedTemplateName = GetStrings().FormatSuggestedTemplateCopyName(state.TemplateName);
                templateName = dialogService.ShowSaveAsTemplateDialog(suggestedTemplateName);
                if (string.IsNullOrWhiteSpace(templateName))
                {
                    return;
                }

                templateCatalog.SaveSheetAsNewTemplate(sheetName, templateName);
                InvalidateRefreshState();
                RefreshActiveTemplateStateFromSheetMetadata();
                TrackTemplateEvent(
                    "ribbon.template.save_as.completed",
                    sheetName,
                    state,
                    properties: new Dictionary<string, object>(StringComparer.Ordinal)
                    {
                        ["templateName"] = templateName,
                    });
                dialogService.ShowInfo(GetStrings().SaveAsTemplateCompletedMessage(templateName));
            }
            catch (Exception ex)
            {
                TrackTemplateEvent(
                    "ribbon.template.save_as.failed",
                    GetOptionalSheetName(),
                    state,
                    properties: new Dictionary<string, object>(StringComparer.Ordinal)
                    {
                        ["templateName"] = templateName ?? string.Empty,
                    },
                    error: CreateTemplateOperationFailedError(ex));
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
                TrackTemplateEvent(
                    "ribbon.template.save.completed",
                    sheetName,
                    state,
                    properties: new Dictionary<string, object>(StringComparer.Ordinal)
                    {
                        ["overwriteRevisionConflict"] = overwriteRevisionConflict,
                    });
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
                ? GetStrings().DefaultTemplateDisplayName
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

        private void TrackTemplateEvent(
            string eventName,
            string sheetName,
            SheetTemplateState state,
            TemplateDefinition template = null,
            AnalyticsError error = null,
            IDictionary<string, object> properties = null)
        {
            if (string.IsNullOrWhiteSpace(eventName))
            {
                return;
            }

            var merged = new Dictionary<string, object>(StringComparer.Ordinal)
            {
                ["sheetName"] = sheetName ?? string.Empty,
                ["projectName"] = state?.ProjectDisplayName ?? template?.ProjectName ?? string.Empty,
                ["templateId"] = template?.TemplateId ?? state?.TemplateId ?? string.Empty,
                ["templateName"] = template?.TemplateName ?? state?.TemplateName ?? string.Empty,
                ["templateRevision"] = template?.Revision ?? state?.TemplateRevision ?? 0,
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

        private string GetOptionalSheetName()
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

        private static AnalyticsError CreateTemplateOperationFailedError(Exception ex)
        {
            return new AnalyticsError
            {
                Code = "template_operation_failed",
                Message = ex?.Message ?? string.Empty,
                ExceptionType = ex?.GetType().Name ?? string.Empty,
            };
        }

        private static HostLocalizedStrings GetStrings()
        {
            return Globals.ThisAddIn?.HostLocalizedStrings ?? HostLocalizedStrings.ForLocale("en");
        }
    }
}
