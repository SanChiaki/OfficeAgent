using System;
using System.Drawing;
using System.Windows.Forms;

namespace OfficeAgent.ExcelAddIn.Dialogs
{
    internal sealed class TemplateNameDialog : Form
    {
        private readonly TextBox templateNameTextBox;

        public TemplateNameDialog(string suggestedTemplateName)
        {
            Text = "另存模板";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(460, 190);

            var root = new TableLayoutPanel
            {
                ColumnCount = 1,
                Dock = DockStyle.Fill,
                Padding = new Padding(16),
                RowCount = 3,
            };
            root.RowStyles.Add(new RowStyle());
            root.RowStyles.Add(new RowStyle());
            root.RowStyles.Add(new RowStyle());

            var promptLabel = new Label
            {
                AutoSize = true,
                Margin = Padding.Empty,
                MaximumSize = new Size(412, 0),
                Text = "请输入新模板名称。保存后，当前表会绑定到新模板。",
            };

            templateNameTextBox = new TextBox
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 12, 0, 12),
                Text = suggestedTemplateName ?? string.Empty,
            };

            var okButton = new Button { Text = "确定", Width = 88, Height = 30 };
            okButton.Click += OkButton_Click;
            var cancelButton = new Button { Text = "取消", Width = 88, Height = 30, DialogResult = DialogResult.Cancel };

            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 46,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(12, 4, 12, 4),
            };
            buttons.Controls.Add(cancelButton);
            buttons.Controls.Add(okButton);

            root.Controls.Add(promptLabel, 0, 0);
            root.Controls.Add(templateNameTextBox, 0, 1);
            root.Controls.Add(buttons, 0, 2);

            Controls.Add(root);
            AcceptButton = okButton;
            CancelButton = cancelButton;
            Shown += (sender, args) =>
            {
                templateNameTextBox.Focus();
                templateNameTextBox.SelectAll();
            };
        }

        public string TemplateName { get; private set; } = string.Empty;

        private void OkButton_Click(object sender, EventArgs e)
        {
            var value = (templateNameTextBox.Text ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(value))
            {
                OperationResultDialog.ShowWarning("模板名称不能为空。");
                return;
            }

            TemplateName = value;
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
