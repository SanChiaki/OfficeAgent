using System;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeAgent.Core.Models;
using OfficeAgent.ExcelAddIn.Localization;

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

        bool ConfirmAiColumnMapping(AiColumnMappingPreview preview);

        AiColumnMappingPreview RunAiColumnMappingWithProgress(
            Func<CancellationToken, Task<AiColumnMappingPreview>> operation);

        SheetBinding ShowProjectLayoutDialog(SheetBinding suggestedBinding);

        void ShowInfo(string message);

        void ShowWarning(string message);

        void ShowError(string message);

        bool ShowAuthenticationRequired(string message);
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

        public bool ConfirmAiColumnMapping(AiColumnMappingPreview preview)
        {
            var owner = ExcelDialogOwner.FromCurrentApplication();
            return AiColumnMappingPreviewDialog.Confirm(preview, owner);
        }

        public AiColumnMappingPreview RunAiColumnMappingWithProgress(
            Func<CancellationToken, Task<AiColumnMappingPreview>> operation)
        {
            var owner = ExcelDialogOwner.FromCurrentApplication();
            return AiColumnMappingProgressDialog.Run(owner, operation);
        }

        public SheetBinding ShowProjectLayoutDialog(SheetBinding suggestedBinding)
        {
            var owner = ExcelDialogOwner.FromCurrentApplication();
            using (var dialog = new ProjectLayoutDialog(suggestedBinding))
            {
                var result = owner == null ? dialog.ShowDialog() : dialog.ShowDialog(owner);
                return result == DialogResult.OK
                    ? dialog.ResultBinding
                    : null;
            }
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

        public bool ShowAuthenticationRequired(string message)
        {
            return OperationResultDialog.ShowAuthenticationRequired(message);
        }
    }

    internal static class OperationResultDialog
    {
        public static void ShowInfo(string message)
        {
            var strings = GetStrings();
            var owner = ExcelDialogOwner.FromCurrentApplication();
            MessageBox.Show(owner, message, strings.HostWindowTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public static void ShowWarning(string message)
        {
            var strings = GetStrings();
            var owner = ExcelDialogOwner.FromCurrentApplication();
            MessageBox.Show(owner, message, strings.HostWindowTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        public static void ShowError(string message)
        {
            var strings = GetStrings();
            var owner = ExcelDialogOwner.FromCurrentApplication();
            MessageBox.Show(owner, message, strings.HostWindowTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        public static bool ShowAuthenticationRequired(string message)
        {
            var owner = ExcelDialogOwner.FromCurrentApplication();
            using (var dialog = new AuthenticationRequiredDialog(message, GetStrings()))
            {
                var result = owner == null ? dialog.ShowDialog() : dialog.ShowDialog(owner);
                return result == DialogResult.Yes;
            }
        }

        private static HostLocalizedStrings GetStrings()
        {
            return Globals.ThisAddIn?.HostLocalizedStrings ?? HostLocalizedStrings.ForLocale("en");
        }

        private sealed class AuthenticationRequiredDialog : Form
        {
            private const int DialogWidth = 360;
            private const int DialogHeight = 140;
            private const int HorizontalPadding = 20;
            private const int ButtonTop = 88;
            private const int ButtonHeight = 28;
            private const int ButtonGap = 8;
            private const int ButtonHorizontalPadding = 18;

            public AuthenticationRequiredDialog(string message, HostLocalizedStrings strings)
            {
                var normalizedMessage = string.IsNullOrWhiteSpace(message)
                    ? strings.AuthenticationRequiredDefaultMessage
                    : message;

                Font = SystemFonts.MessageBoxFont;
                AutoScaleMode = AutoScaleMode.Dpi;
                Text = strings.HostWindowTitle;
                StartPosition = FormStartPosition.CenterScreen;
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false;
                MinimizeBox = false;
                ShowInTaskbar = false;
                ClientSize = new Size(DialogWidth, DialogHeight);

                var messageLabel = new Label
                {
                    AutoSize = false,
                    Text = normalizedMessage,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Bounds = new Rectangle(
                        HorizontalPadding,
                        20,
                        DialogWidth - (HorizontalPadding * 2),
                        44),
                };

                var closeButtonWidth = MeasureButtonWidth(strings.CloseButtonText);
                var loginButtonWidth = MeasureButtonWidth(strings.AuthenticationRequiredLoginButtonText);
                var closeButtonLeft = DialogWidth - HorizontalPadding - closeButtonWidth;
                var loginButtonLeft = closeButtonLeft - ButtonGap - loginButtonWidth;

                var loginButton = new Button
                {
                    Text = strings.AuthenticationRequiredLoginButtonText,
                    DialogResult = DialogResult.Yes,
                    Bounds = new Rectangle(loginButtonLeft, ButtonTop, loginButtonWidth, ButtonHeight),
                };

                var closeButton = new Button
                {
                    Text = strings.CloseButtonText,
                    DialogResult = DialogResult.Cancel,
                    Bounds = new Rectangle(closeButtonLeft, ButtonTop, closeButtonWidth, ButtonHeight),
                };

                AcceptButton = loginButton;
                CancelButton = closeButton;

                Controls.Add(messageLabel);
                Controls.Add(loginButton);
                Controls.Add(closeButton);
            }

            private int MeasureButtonWidth(string text)
            {
                var measured = TextRenderer.MeasureText(
                    text ?? string.Empty,
                    Font,
                    new Size(int.MaxValue, ButtonHeight),
                    TextFormatFlags.SingleLine | TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);

                return Math.Max(72, measured.Width + ButtonHorizontalPadding);
            }
        }
    }
}
