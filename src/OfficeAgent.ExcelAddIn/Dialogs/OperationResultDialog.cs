using System.Windows.Forms;
using OfficeAgent.Core.Models;

namespace OfficeAgent.ExcelAddIn.Dialogs
{
    public interface IRibbonSyncDialogService
    {
        bool ConfirmDownload(
            string operationName,
            string projectName,
            int rowCount,
            int fieldCount,
            SyncOperationPreview overwritePreview);

        bool ConfirmUpload(string operationName, string projectName, SyncOperationPreview preview);

        void ShowInfo(string message);

        void ShowWarning(string message);

        void ShowError(string message);
    }

    internal sealed class RibbonSyncDialogService : IRibbonSyncDialogService
    {
        public bool ConfirmDownload(
            string operationName,
            string projectName,
            int rowCount,
            int fieldCount,
            SyncOperationPreview overwritePreview)
        {
            return DownloadConfirmDialog.Confirm(operationName, projectName, rowCount, fieldCount, overwritePreview);
        }

        public bool ConfirmUpload(string operationName, string projectName, SyncOperationPreview preview)
        {
            return UploadConfirmDialog.Confirm(operationName, projectName, preview);
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

    internal static class OperationResultDialog
    {
        public static void ShowInfo(string message)
        {
            MessageBox.Show(message, "Resy AI", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public static void ShowWarning(string message)
        {
            MessageBox.Show(message, "Resy AI", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        public static void ShowError(string message)
        {
            MessageBox.Show(message, "Resy AI", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
