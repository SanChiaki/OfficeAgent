using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using OfficeAgent.Core.Models;

namespace OfficeAgent.ExcelAddIn.Dialogs
{
    internal sealed class TemplatePickerDialog : Form
    {
        private readonly ListBox templateListBox;

        public TemplatePickerDialog(string projectDisplayName, IReadOnlyList<TemplateDefinition> templates)
        {
            Text = "应用模板";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(560, 360);

            var root = new TableLayoutPanel
            {
                ColumnCount = 1,
                Dock = DockStyle.Fill,
                Padding = new Padding(16),
                RowCount = 3,
            };
            root.RowStyles.Add(new RowStyle());
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle());

            var headerPanel = new TableLayoutPanel
            {
                AutoSize = true,
                ColumnCount = 1,
                Dock = DockStyle.Fill,
                Margin = Padding.Empty,
                RowCount = 2,
            };
            headerPanel.RowStyles.Add(new RowStyle());
            headerPanel.RowStyles.Add(new RowStyle());

            var projectLabel = new Label
            {
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 6),
                MaximumSize = new Size(496, 0),
                Text = $"当前项目：{projectDisplayName}",
            };

            var instructionLabel = new Label
            {
                AutoSize = true,
                Margin = Padding.Empty,
                MaximumSize = new Size(496, 0),
                Text = "请选择要应用到当前表的本机模板。",
            };

            headerPanel.Controls.Add(projectLabel, 0, 0);
            headerPanel.Controls.Add(instructionLabel, 0, 1);

            templateListBox = new ListBox
            {
                Dock = DockStyle.Fill,
                HorizontalScrollbar = true,
                IntegralHeight = false,
                Margin = new Padding(0, 12, 0, 12),
            };

            foreach (var template in (templates ?? Array.Empty<TemplateDefinition>()).OrderByDescending(item => item.UpdatedAtUtc))
            {
                templateListBox.Items.Add(new TemplateListItem(template));
            }

            if (templateListBox.Items.Count > 0)
            {
                templateListBox.SelectedIndex = 0;
            }

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

            root.Controls.Add(headerPanel, 0, 0);
            root.Controls.Add(templateListBox, 0, 1);
            root.Controls.Add(buttons, 0, 2);

            Controls.Add(root);
            AcceptButton = okButton;
            CancelButton = cancelButton;
        }

        public string SelectedTemplateId { get; private set; } = string.Empty;

        private void OkButton_Click(object sender, EventArgs e)
        {
            var selected = templateListBox.SelectedItem as TemplateListItem;
            if (selected == null)
            {
                OperationResultDialog.ShowWarning("请选择一个模板。");
                return;
            }

            SelectedTemplateId = selected.TemplateId;
            DialogResult = DialogResult.OK;
            Close();
        }

        private sealed class TemplateListItem
        {
            public TemplateListItem(TemplateDefinition template)
            {
                TemplateId = template?.TemplateId ?? string.Empty;
                DisplayText = string.Format(
                    "{0}  (v{1})",
                    template?.TemplateName ?? string.Empty,
                    template?.Revision ?? 0);
            }

            public string TemplateId { get; }

            public string DisplayText { get; }

            public override string ToString()
            {
                return DisplayText;
            }
        }
    }
}
