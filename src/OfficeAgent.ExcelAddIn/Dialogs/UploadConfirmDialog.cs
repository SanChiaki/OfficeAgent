using System.Text;
using OfficeAgent.Core.Models;
using System.Windows.Forms;

namespace OfficeAgent.ExcelAddIn.Dialogs
{
    internal static class UploadConfirmDialog
    {
        public static bool Confirm(string operationName, string projectName, SyncOperationPreview preview)
        {
            var builder = new StringBuilder()
                .AppendLine($"确认要执行{operationName}吗？")
                .AppendLine($"项目：{projectName}")
                .AppendLine(preview?.Summary ?? string.Empty);

            foreach (var detail in preview?.Details ?? System.Array.Empty<string>())
            {
                builder.AppendLine(detail);
            }

            var result = MessageBox.Show(
                builder.ToString(),
                "Resy AI",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2);
            return result == DialogResult.Yes;
        }
    }
}
