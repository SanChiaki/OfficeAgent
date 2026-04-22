using System.Windows.Forms;

namespace OfficeAgent.ExcelAddIn.Dialogs
{
    internal static class TemplateRevisionConflictDialog
    {
        public static DialogResult ShowDecision(string templateName, int sheetRevision, int storedRevision)
        {
            return TemplatePromptDialog.ShowPrompt(
                "模板版本冲突",
                $"模板“{templateName}”已从版本 {sheetRevision} 更新到版本 {storedRevision}。\r\n请选择后续操作。",
                MessageBoxIcon.Warning,
                new TemplatePromptDialog.DialogButtonSpec("取消", DialogResult.Cancel, isCancel: true),
                new TemplatePromptDialog.DialogButtonSpec("另存为新模板", DialogResult.No),
                new TemplatePromptDialog.DialogButtonSpec("覆盖原模板", DialogResult.Yes, isAccept: true));
        }
    }
}
