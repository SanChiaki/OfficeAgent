using System.Text;
using OfficeAgent.Core.Models;
using System.Windows.Forms;

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
            var builder = new StringBuilder()
                .AppendLine($"确认要执行{operationName}吗？")
                .AppendLine($"项目：{projectName}")
                .AppendLine($"记录数：{rowCount}")
                .AppendLine($"字段数：{fieldCount}");

            var dirtyCount = overwritePreview?.Changes?.Length ?? 0;
            if (dirtyCount > 0)
            {
                builder
                    .AppendLine()
                    .AppendLine($"将覆盖 {dirtyCount} 个未上传改单元格。");

                foreach (var detail in overwritePreview.Details ?? System.Array.Empty<string>())
                {
                    builder.AppendLine(detail);
                }
            }

            var result = MessageBox.Show(
                builder.ToString(),
                "Resy AI",
                MessageBoxButtons.YesNo,
                dirtyCount > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Question,
                dirtyCount > 0 ? MessageBoxDefaultButton.Button2 : MessageBoxDefaultButton.Button1);
            return result == DialogResult.Yes;
        }
    }
}
