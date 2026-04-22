using System;
using System.Collections.Generic;
using System.Windows.Forms;
using OfficeAgent.Core.Models;

namespace OfficeAgent.ExcelAddIn.Dialogs
{
    public interface IRibbonTemplateDialogService
    {
        string ShowTemplatePicker(string projectDisplayName, IReadOnlyList<TemplateDefinition> templates);

        string ShowSaveAsTemplateDialog(string suggestedTemplateName);

        bool ConfirmApplyTemplateOverwrite(string templateName);

        DialogResult ShowTemplateRevisionConflictDialog(string templateName, int sheetRevision, int storedRevision);

        void ShowInfo(string message);

        void ShowWarning(string message);

        void ShowError(string message);
    }

    internal sealed class RibbonTemplateDialogService : IRibbonTemplateDialogService
    {
        public string ShowTemplatePicker(string projectDisplayName, IReadOnlyList<TemplateDefinition> templates)
        {
            using (var dialog = new TemplatePickerDialog(projectDisplayName, templates))
            {
                return dialog.ShowDialog() == DialogResult.OK
                    ? dialog.SelectedTemplateId
                    : string.Empty;
            }
        }

        public string ShowSaveAsTemplateDialog(string suggestedTemplateName)
        {
            using (var dialog = new TemplateNameDialog(suggestedTemplateName))
            {
                return dialog.ShowDialog() == DialogResult.OK
                    ? dialog.TemplateName
                    : string.Empty;
            }
        }

        public bool ConfirmApplyTemplateOverwrite(string templateName)
        {
            return TemplateOverwriteConfirmDialog.Confirm(templateName);
        }

        public DialogResult ShowTemplateRevisionConflictDialog(string templateName, int sheetRevision, int storedRevision)
        {
            return TemplateRevisionConflictDialog.ShowDecision(templateName, sheetRevision, storedRevision);
        }

        public void ShowInfo(string message)
        {
            OperationResultDialog.ShowInfo(message);
        }

        public void ShowWarning(string message)
        {
            OperationResultDialog.ShowWarning(message);
        }

        public void ShowError(string message)
        {
            OperationResultDialog.ShowError(message);
        }
    }
}
