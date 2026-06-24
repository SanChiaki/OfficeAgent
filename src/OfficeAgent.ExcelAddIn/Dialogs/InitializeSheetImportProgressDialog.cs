using System;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeAgent.ExcelAddIn.Localization;

namespace OfficeAgent.ExcelAddIn.Dialogs
{
    internal interface IInitializeSheetImportProgress
    {
        void SetDownloading();

        void SetImporting();

        void SetWritingConfiguration();
    }

    internal sealed class InitializeSheetImportProgressDialog : Form, IInitializeSheetImportProgress
    {
        private readonly Func<IInitializeSheetImportProgress, CancellationToken, Task> operation;
        private readonly HostLocalizedStrings strings;
        private readonly CancellationTokenSource cancellationTokenSource = new CancellationTokenSource();
        private readonly Label messageLabel;
        private readonly Button cancelButton;
        private Exception error;
        private bool completed;
        private bool canceled;

        private InitializeSheetImportProgressDialog(
            Func<IInitializeSheetImportProgress, CancellationToken, Task> operation,
            HostLocalizedStrings strings)
        {
            this.operation = operation ?? throw new ArgumentNullException(nameof(operation));
            this.strings = strings ?? HostLocalizedStrings.ForLocale("en");

            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Dpi;
            Text = this.strings.InitializeSheetImportProgressDialogTitle;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            ControlBox = false;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(400, 150);
            Padding = new Padding(18);

            messageLabel = new Label
            {
                AutoSize = false,
                Dock = DockStyle.Top,
                Height = 58,
                Text = this.strings.InitializeSheetImportProgressDownloadingText,
                TextAlign = ContentAlignment.MiddleLeft,
            };

            var progressBar = new ProgressBar
            {
                Dock = DockStyle.Top,
                Height = 18,
                MarqueeAnimationSpeed = 30,
                Style = ProgressBarStyle.Marquee,
            };

            cancelButton = new Button
            {
                AutoSize = true,
                Padding = new Padding(14, 4, 14, 4),
                Text = this.strings.CancelButtonText,
            };
            cancelButton.Click += (sender, args) =>
            {
                cancelButton.Enabled = false;
                cancellationTokenSource.Cancel();
            };

            var buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                FlowDirection = FlowDirection.RightToLeft,
                Height = 42,
                Padding = new Padding(0, 12, 0, 0),
            };
            buttonPanel.Controls.Add(cancelButton);

            Controls.Add(progressBar);
            Controls.Add(messageLabel);
            Controls.Add(buttonPanel);

            SetDownloading();
        }

        public static bool Run(
            IWin32Window owner,
            Func<IInitializeSheetImportProgress, CancellationToken, Task> operation)
        {
            var strings = Globals.ThisAddIn?.HostLocalizedStrings ?? HostLocalizedStrings.ForLocale("en");
            using (var dialog = new InitializeSheetImportProgressDialog(operation, strings))
            {
                var dialogResult = owner == null
                    ? dialog.ShowDialog()
                    : dialog.ShowDialog(owner);

                if (dialog.error != null)
                {
                    throw dialog.error;
                }

                return dialogResult == DialogResult.OK && dialog.completed && !dialog.canceled;
            }
        }

        public void SetDownloading()
        {
            SetProgressText(strings.InitializeSheetImportProgressDownloadingText, cancelEnabled: true);
        }

        public void SetImporting()
        {
            SetProgressText(strings.InitializeSheetImportProgressImportingText, cancelEnabled: false);
        }

        public void SetWritingConfiguration()
        {
            SetProgressText(strings.InitializeSheetImportProgressWritingConfigurationText, cancelEnabled: false);
        }

        protected override async void OnShown(EventArgs e)
        {
            base.OnShown(e);

            try
            {
                await operation(this, cancellationTokenSource.Token);
                completed = true;
                DialogResult = DialogResult.OK;
            }
            catch (OperationCanceledException)
            {
                canceled = true;
                DialogResult = DialogResult.Cancel;
            }
            catch (Exception ex)
            {
                error = ex;
                DialogResult = DialogResult.Abort;
            }
            finally
            {
                Close();
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (!completed)
                {
                    cancellationTokenSource.Cancel();
                }

                cancellationTokenSource.Dispose();
            }

            base.Dispose(disposing);
        }

        private void SetProgressText(string message, bool cancelEnabled)
        {
            RunOnUiThread(() =>
            {
                messageLabel.Text = message ?? string.Empty;
                cancelButton.Enabled = cancelEnabled && !cancellationTokenSource.IsCancellationRequested;
            });
        }

        private void RunOnUiThread(Action action)
        {
            if (action == null)
            {
                throw new ArgumentNullException(nameof(action));
            }

            if (IsDisposed)
            {
                return;
            }

            if (InvokeRequired)
            {
                BeginInvoke(action);
                return;
            }

            action();
        }
    }
}
