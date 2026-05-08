using System;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeAgent.Core.Models;
using OfficeAgent.ExcelAddIn.Localization;

namespace OfficeAgent.ExcelAddIn.Dialogs
{
    internal sealed class AiColumnMappingProgressDialog : Form
    {
        private readonly Func<CancellationToken, Task<AiColumnMappingPreview>> operation;
        private readonly CancellationTokenSource cancellationTokenSource = new CancellationTokenSource();
        private AiColumnMappingPreview result;
        private Exception error;
        private bool completed;

        private AiColumnMappingProgressDialog(
            Func<CancellationToken, Task<AiColumnMappingPreview>> operation,
            HostLocalizedStrings strings)
        {
            this.operation = operation ?? throw new ArgumentNullException(nameof(operation));
            strings = strings ?? HostLocalizedStrings.ForLocale("en");

            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Dpi;
            Text = strings.AiColumnMappingProgressDialogTitle;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            ControlBox = false;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(380, 150);
            Padding = new Padding(18);

            var messageLabel = new Label
            {
                AutoSize = false,
                Dock = DockStyle.Top,
                Height = 58,
                Text = strings.AiColumnMappingProgressMessage,
                TextAlign = ContentAlignment.MiddleLeft,
            };

            var progressBar = new ProgressBar
            {
                Dock = DockStyle.Top,
                Height = 18,
                MarqueeAnimationSpeed = 30,
                Style = ProgressBarStyle.Marquee,
            };

            var cancelButton = new Button
            {
                Text = strings.AiColumnMappingAbortButtonText,
                AutoSize = true,
                Padding = new Padding(14, 4, 14, 4),
                Anchor = AnchorStyles.Right,
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
        }

        public static AiColumnMappingPreview Run(
            IWin32Window owner,
            Func<CancellationToken, Task<AiColumnMappingPreview>> operation)
        {
            var strings = Globals.ThisAddIn?.HostLocalizedStrings ?? HostLocalizedStrings.ForLocale("en");
            using (var dialog = new AiColumnMappingProgressDialog(operation, strings))
            {
                var dialogResult = owner == null
                    ? dialog.ShowDialog()
                    : dialog.ShowDialog(owner);
                if (dialog.error != null)
                {
                    throw dialog.error;
                }

                return dialogResult == DialogResult.OK && dialog.completed
                    ? dialog.result
                    : null;
            }
        }

        protected override async void OnShown(EventArgs e)
        {
            base.OnShown(e);

            try
            {
                result = await operation(cancellationTokenSource.Token);
                completed = true;
                DialogResult = DialogResult.OK;
            }
            catch (OperationCanceledException)
            {
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
    }
}
