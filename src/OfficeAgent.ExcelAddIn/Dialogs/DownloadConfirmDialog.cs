using System.Text;
using OfficeAgent.Core.Models;
using System.Windows.Forms;
using OfficeAgent.ExcelAddIn.Localization;

namespace OfficeAgent.ExcelAddIn.Dialogs
{
    internal static class DownloadConfirmDialog
    {
        public static bool Confirm(
            string operationName,
            string projectName,
            int rowCount,
            int fieldCount,
            SyncOperationPreview overwritePreview)
        {
            var strings = Globals.ThisAddIn?.HostLocalizedStrings ?? HostLocalizedStrings.ForLocale("en");
            var builder = new StringBuilder()
                .AppendLine(strings.ConfirmOperationPrompt(operationName))
                .AppendLine(strings.ProjectLine(projectName))
                .AppendLine(strings.RowCountLine(rowCount))
                .AppendLine(strings.FieldCountLine(fieldCount));

            var dirtyCount = overwritePreview?.Changes?.Length ?? 0;
            if (dirtyCount > 0)
            {
                builder
                    .AppendLine()
                    .AppendLine(strings.OverwriteDirtyCellsLine(dirtyCount));

                foreach (var detail in overwritePreview.Details ?? System.Array.Empty<string>())
                {
                    builder.AppendLine(detail);
                }
            }

            var result = TemplatePromptDialog.ShowPrompt(
                strings.HostWindowTitle,
                builder.ToString(),
                dirtyCount > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Question,
                new TemplatePromptDialog.DialogButtonSpec(strings.NoButtonText, DialogResult.No, isCancel: true),
                new TemplatePromptDialog.DialogButtonSpec(strings.YesButtonText, DialogResult.Yes, isAccept: true));
            return result == DialogResult.Yes;
        }
    }
}
