using System;
using System.Drawing;
using System.Windows.Forms;
using OfficeAgent.Core.Models;

namespace OfficeAgent.ExcelAddIn.Dialogs
{
    internal sealed class ProjectLayoutDialog : Form
    {
        private readonly TextBox headerStartRowTextBox;
        private readonly TextBox headerRowCountTextBox;
        private readonly TextBox dataStartRowTextBox;
        private readonly SheetBinding suggestedBinding;

        public ProjectLayoutDialog(SheetBinding suggestedBinding)
        {
            this.suggestedBinding = suggestedBinding ?? throw new ArgumentNullException(nameof(suggestedBinding));

            Text = "配置当前表布局";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(520, 260);

            var instructionLabel = new Label
            {
                AutoSize = false,
                Location = new Point(16, 12),
                Size = new Size(488, 44),
                Text = "下面三个值会写入当前工作表的同步配置（SheetBindings），请确认后保存。",
            };

            var projectLabel = new Label
            {
                AutoSize = false,
                Location = new Point(16, 60),
                Size = new Size(488, 20),
                Text = FormatProjectLabel(suggestedBinding),
            };

            var headerStartRowLabel = new Label
            {
                AutoSize = true,
                Location = new Point(16, 95),
                Text = "HeaderStartRow",
            };
            headerStartRowTextBox = CreateValueTextBox("HeaderStartRowTextBox", 16, 116, suggestedBinding.HeaderStartRow);

            var headerRowCountLabel = new Label
            {
                AutoSize = true,
                Location = new Point(184, 95),
                Text = "HeaderRowCount",
            };
            headerRowCountTextBox = CreateValueTextBox("HeaderRowCountTextBox", 184, 116, suggestedBinding.HeaderRowCount);

            var dataStartRowLabel = new Label
            {
                AutoSize = true,
                Location = new Point(352, 95),
                Text = "DataStartRow",
            };
            dataStartRowTextBox = CreateValueTextBox("DataStartRowTextBox", 352, 116, suggestedBinding.DataStartRow);

            var okButton = new Button
            {
                Text = "确定",
                DialogResult = DialogResult.None,
                Location = new Point(336, 210),
                Size = new Size(80, 28),
            };
            okButton.Click += HandleOkClick;

            var cancelButton = new Button
            {
                Text = "取消",
                DialogResult = DialogResult.Cancel,
                Location = new Point(424, 210),
                Size = new Size(80, 28),
            };

            AcceptButton = okButton;
            CancelButton = cancelButton;

            Controls.Add(instructionLabel);
            Controls.Add(projectLabel);
            Controls.Add(headerStartRowLabel);
            Controls.Add(headerStartRowTextBox);
            Controls.Add(headerRowCountLabel);
            Controls.Add(headerRowCountTextBox);
            Controls.Add(dataStartRowLabel);
            Controls.Add(dataStartRowTextBox);
            Controls.Add(okButton);
            Controls.Add(cancelButton);
        }

        public SheetBinding ResultBinding { get; private set; }

        private void HandleOkClick(object sender, EventArgs e)
        {
            if (!TryCreateBinding(
                suggestedBinding,
                headerStartRowTextBox.Text,
                headerRowCountTextBox.Text,
                dataStartRowTextBox.Text,
                out var binding,
                out var errorMessage))
            {
                MessageBox.Show(this, errorMessage, "Resy AI", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            ResultBinding = binding;
            DialogResult = DialogResult.OK;
            Close();
        }

        private static TextBox CreateValueTextBox(string name, int left, int top, int value)
        {
            return new TextBox
            {
                Name = name,
                Location = new Point(left, top),
                Size = new Size(152, 23),
                Text = value.ToString(),
            };
        }

        private static string FormatProjectLabel(SheetBinding binding)
        {
            return string.Format("当前绑定：{0} | {1}", binding.ProjectId, binding.ProjectName);
        }

        private static bool TryCreateBinding(
            SheetBinding suggestedBinding,
            string headerStartRowText,
            string headerRowCountText,
            string dataStartRowText,
            out SheetBinding binding,
            out string errorMessage)
        {
            if (suggestedBinding == null)
            {
                throw new ArgumentNullException(nameof(suggestedBinding));
            }

            binding = null;
            errorMessage = null;

            if (!TryParsePositiveInt(headerStartRowText, out var headerStartRow))
            {
                errorMessage = "HeaderStartRow 必须是正整数。";
                return false;
            }

            if (!TryParsePositiveInt(headerRowCountText, out var headerRowCount))
            {
                errorMessage = "HeaderRowCount 必须是正整数。";
                return false;
            }

            if (!TryParsePositiveInt(dataStartRowText, out var dataStartRow))
            {
                errorMessage = "DataStartRow 必须是正整数。";
                return false;
            }

            if (dataStartRow < headerStartRow + headerRowCount)
            {
                errorMessage = "DataStartRow 必须大于或等于 HeaderStartRow + HeaderRowCount。";
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
