using System.Windows.Forms;

namespace OfficeAgent.ExcelAddIn.Dialogs
{
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
