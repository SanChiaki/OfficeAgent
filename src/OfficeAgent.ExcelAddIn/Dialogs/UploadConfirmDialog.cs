using System.Text;
using OfficeAgent.Core.Models;
using System.Windows.Forms;
using OfficeAgent.ExcelAddIn.Localization;

namespace OfficeAgent.ExcelAddIn.Dialogs
{
    internal static class UploadConfirmDialog
    {
        public static bool Confirm(string operationName, string projectName, SyncOperationPreview preview)
        {
            var strings = Globals.ThisAddIn?.HostLocalizedStrings ?? HostLocalizedStrings.ForLocale("en");
            var builder = new StringBuilder()
                .AppendLine(strings.ConfirmOperationPrompt(operationName))
                .AppendLine(strings.ProjectLine(projectName))
                .AppendLine(preview?.Summary ?? string.Empty);

            foreach (var detail in preview?.Details ?? System.Array.Empty<string>())
            {
                builder.AppendLine(detail);
            }

            var result = TemplatePromptDialog.ShowPrompt(
                strings.HostWindowTitle,
                builder.ToString(),
                MessageBoxIcon.Question,
                new TemplatePromptDialog.DialogButtonSpec(strings.NoButtonText, DialogResult.No, isCancel: true),
                new TemplatePromptDialog.DialogButtonSpec(strings.YesButtonText, DialogResult.Yes, isAccept: true));
            return result == DialogResult.Yes;
        }
    }
}
