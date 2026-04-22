using System.Windows.Forms;

namespace OfficeAgent.ExcelAddIn.Dialogs
{
    internal static class TemplateOverwriteConfirmDialog
    {
        public static bool Confirm(string templateName)
        {
            var result = TemplatePromptDialog.ShowPrompt(
                "覆盖模板",
                $"当前表存在未保存的模板改动，确认用模板“{templateName}”覆盖吗？",
                MessageBoxIcon.Warning,
                new TemplatePromptDialog.DialogButtonSpec("取消", DialogResult.No, isCancel: true),
                new TemplatePromptDialog.DialogButtonSpec("覆盖", DialogResult.Yes, isAccept: true));

            return result == DialogResult.Yes;
        }
    }
}
