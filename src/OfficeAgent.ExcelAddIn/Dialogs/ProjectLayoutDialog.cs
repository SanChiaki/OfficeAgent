using System;
using System.Drawing;
using System.Windows.Forms;
using OfficeAgent.Core.Diagnostics;
using OfficeAgent.Core.Models;
using OfficeAgent.ExcelAddIn.Localization;

namespace OfficeAgent.ExcelAddIn.Dialogs
{
    internal sealed class ProjectLayoutDialog : Form
    {
        private readonly TextBox headerStartRowTextBox;
        private readonly TextBox headerRowCountTextBox;
        private readonly TextBox dataStartRowTextBox;
        private readonly SheetBinding suggestedBinding;
        private readonly HostLocalizedStrings strings;

        public ProjectLayoutDialog(SheetBinding suggestedBinding, HostLocalizedStrings strings = null)
        {
            this.suggestedBinding = suggestedBinding ?? throw new ArgumentNullException(nameof(suggestedBinding));
            this.strings = strings ?? Globals.ThisAddIn?.HostLocalizedStrings ?? HostLocalizedStrings.ForLocale("en");

            Text = this.strings.ProjectLayoutDialogTitle;
            StartPosition = FormStartPosition.CenterParent;
            AutoScaleMode = AutoScaleMode.Font;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            Padding = new Padding(16);

            var instructionLabel = new Label
            {
                AutoSize = true,
                Margin = new Padding(0),
                MaximumSize = new Size(520, 0),
                Text = this.strings.ProjectLayoutInstructionText,
            };

            var projectLabel = new Label
            {
                AutoSize = true,
                Margin = new Padding(0, 8, 0, 0),
                MaximumSize = new Size(520, 0),
                Text = FormatProjectLabel(suggestedBinding, this.strings),
            };

            var fieldsLayout = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Margin = new Padding(0, 16, 0, 0),
                WrapContents = false,
            };
            fieldsLayout.Controls.Add(CreateFieldEditor(
                "HeaderStartRow",
                "HeaderStartRowTextBox",
                suggestedBinding.HeaderStartRow,
                new Padding(0, 0, 16, 0),
                out headerStartRowTextBox));
            fieldsLayout.Controls.Add(CreateFieldEditor(
                "HeaderRowCount",
                "HeaderRowCountTextBox",
                suggestedBinding.HeaderRowCount,
                new Padding(0, 0, 16, 0),
                out headerRowCountTextBox));
            fieldsLayout.Controls.Add(CreateFieldEditor(
                "DataStartRow",
                "DataStartRowTextBox",
                suggestedBinding.DataStartRow,
                new Padding(0),
                out dataStartRowTextBox));

            var okButton = new Button
            {
                Text = this.strings.OkButtonText,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                DialogResult = DialogResult.None,
                Margin = new Padding(8, 0, 0, 0),
                Padding = new Padding(10, 4, 10, 4),
            };
            okButton.Click += HandleOkClick;

            var cancelButton = new Button
            {
                Text = this.strings.CancelButtonText,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                DialogResult = DialogResult.Cancel,
                Margin = new Padding(8, 0, 0, 0),
                Padding = new Padding(10, 4, 10, 4),
            };

            var buttonsLayout = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                FlowDirection = FlowDirection.RightToLeft,
                Margin = new Padding(0, 16, 0, 0),
                WrapContents = false,
            };
            buttonsLayout.Controls.Add(cancelButton);
            buttonsLayout.Controls.Add(okButton);

            var contentLayout = new TableLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 1,
                Dock = DockStyle.Fill,
                Margin = new Padding(0),
                Padding = new Padding(0),
            };
            contentLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            AcceptButton = okButton;
            CancelButton = cancelButton;

            contentLayout.Controls.Add(instructionLabel);
            contentLayout.Controls.Add(projectLabel);
            contentLayout.Controls.Add(fieldsLayout);
            contentLayout.Controls.Add(buttonsLayout);

            Controls.Add(contentLayout);
        }

        public SheetBinding ResultBinding { get; private set; }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            OfficeAgentLog.Info(
                "ribbon_sync",
                "project.layout_dialog.closed",
                "Project layout dialog closed.",
                $"DialogResult={DialogResult}; CloseReason={e?.CloseReason}; SheetName={suggestedBinding.SheetName ?? string.Empty}; ProjectId={suggestedBinding.ProjectId ?? string.Empty}; ProjectName={suggestedBinding.ProjectName ?? string.Empty}; HeaderStartRow={suggestedBinding.HeaderStartRow}; HeaderRowCount={suggestedBinding.HeaderRowCount}; DataStartRow={suggestedBinding.DataStartRow}");

            base.OnFormClosed(e);
        }

        private void HandleOkClick(object sender, EventArgs e)
        {
            if (!TryCreateBinding(
                suggestedBinding,
                headerStartRowTextBox.Text,
                headerRowCountTextBox.Text,
                dataStartRowTextBox.Text,
                strings,
                out var binding,
                out var errorMessage))
            {
                MessageBox.Show(this, errorMessage, strings.HostWindowTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            ResultBinding = binding;
            DialogResult = DialogResult.OK;
            Close();
        }

        private static Control CreateFieldEditor(string labelText, string textBoxName, int value, Padding margin, out TextBox textBox)
        {
            var fieldLayout = new TableLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 1,
                Margin = margin,
                Padding = new Padding(0),
            };
            fieldLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            fieldLayout.Controls.Add(new Label
            {
                AutoSize = true,
                Margin = new Padding(0),
                Text = labelText,
            });

            textBox = CreateValueTextBox(textBoxName, value);
            fieldLayout.Controls.Add(textBox);
            return fieldLayout;
        }

        private static TextBox CreateValueTextBox(string name, int value)
        {
            var textBox = new TextBox
            {
                Name = name,
                Margin = new Padding(0, 8, 0, 0),
                Text = value.ToString(),
            };
            textBox.Width = 152;
            return textBox;
        }

        private static string FormatProjectLabel(SheetBinding binding, HostLocalizedStrings strings)
        {
            return strings.ProjectLayoutCurrentBindingText(binding.ProjectId, binding.ProjectName);
        }

        private static bool TryCreateBinding(
            SheetBinding suggestedBinding,
            string headerStartRowText,
            string headerRowCountText,
            string dataStartRowText,
            HostLocalizedStrings strings,
            out SheetBinding binding,
            out string errorMessage)
        {
            if (suggestedBinding == null)
            {
                throw new ArgumentNullException(nameof(suggestedBinding));
            }

            binding = null;
            errorMessage = null;
            var localizedStrings = strings ?? HostLocalizedStrings.ForLocale("en");

            if (!TryParsePositiveInt(headerStartRowText, out var headerStartRow))
            {
                errorMessage = localizedStrings.ProjectLayoutPositiveIntegerError("HeaderStartRow");
                return false;
            }

            if (!TryParsePositiveInt(headerRowCountText, out var headerRowCount))
            {
                errorMessage = localizedStrings.ProjectLayoutPositiveIntegerError("HeaderRowCount");
                return false;
            }

            if (!TryParsePositiveInt(dataStartRowText, out var dataStartRow))
            {
                errorMessage = localizedStrings.ProjectLayoutPositiveIntegerError("DataStartRow");
                return false;
            }

            if (dataStartRow < headerStartRow + headerRowCount)
            {
                errorMessage = localizedStrings.ProjectLayoutDataStartValidationError;
                return false;
            }

            binding = new SheetBinding
            {
                SheetName = suggestedBinding.SheetName,
                SystemKey = suggestedBinding.SystemKey,
                ProjectId = suggestedBinding.ProjectId,
                ProjectName = suggestedBinding.ProjectName,
                HeaderStartRow = headerStartRow,
                HeaderRowCount = headerRowCount,
                DataStartRow = dataStartRow,
            };
            return true;
        }

        private static bool TryParsePositiveInt(string text, out int value)
        {
            return int.TryParse(text, out value) && value > 0;
        }
    }
}
