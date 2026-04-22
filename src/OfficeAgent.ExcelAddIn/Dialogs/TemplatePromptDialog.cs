using System;
using System.Drawing;
using System.Windows.Forms;

namespace OfficeAgent.ExcelAddIn.Dialogs
{
    internal sealed class TemplatePromptDialog : Form
    {
        private TemplatePromptDialog(
            string title,
            string message,
            MessageBoxIcon icon,
            params DialogButtonSpec[] buttons)
        {
            if (buttons == null || buttons.Length == 0)
            {
                throw new ArgumentException("At least one button is required.", nameof(buttons));
            }

            Text = title ?? "ISDP";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(540, 220);

            var root = new TableLayoutPanel
            {
                ColumnCount = 1,
                Dock = DockStyle.Fill,
                Padding = new Padding(16),
                RowCount = 2,
            };
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle());

            var content = new TableLayoutPanel
            {
                ColumnCount = 2,
                Dock = DockStyle.Fill,
                RowCount = 1,
            };
            content.ColumnStyles.Add(new ColumnStyle());
            content.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

            var iconBox = new PictureBox
            {
                Image = ResolveIcon(icon).ToBitmap(),
                Margin = new Padding(0, 4, 16, 0),
                Size = new Size(32, 32),
                SizeMode = PictureBoxSizeMode.StretchImage,
            };

            var messageLabel = new Label
            {
                AutoSize = true,
                Dock = DockStyle.Fill,
                Margin = Padding.Empty,
                MaximumSize = new Size(440, 0),
                Text = message ?? string.Empty,
            };

            content.Controls.Add(iconBox, 0, 0);
            content.Controls.Add(messageLabel, 1, 0);

            var buttonsPanel = new FlowLayoutPanel
            {
                AutoSize = true,
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                Margin = new Padding(0, 16, 0, 0),
                Padding = Padding.Empty,
                WrapContents = false,
            };

            foreach (var spec in buttons)
            {
                var button = new Button
                {
                    DialogResult = spec.Result,
                    Height = 30,
                    MinimumSize = new Size(88, 30),
                    Text = spec.Text ?? string.Empty,
                    Width = Math.Max(88, TextRenderer.MeasureText(spec.Text ?? string.Empty, SystemFonts.MessageBoxFont).Width + 24),
                };

                if (spec.IsAccept)
                {
                    AcceptButton = button;
                }

                if (spec.IsCancel)
                {
                    CancelButton = button;
                }

                buttonsPanel.Controls.Add(button);
            }

            root.Controls.Add(content, 0, 0);
            root.Controls.Add(buttonsPanel, 0, 1);
            Controls.Add(root);
        }

        public static DialogResult ShowPrompt(
            string title,
            string message,
            MessageBoxIcon icon,
            params DialogButtonSpec[] buttons)
        {
            using (var dialog = new TemplatePromptDialog(title, message, icon, buttons))
            {
                return dialog.ShowDialog();
            }
        }

        private static Icon ResolveIcon(MessageBoxIcon icon)
        {
            switch (icon)
            {
                case MessageBoxIcon.Error:
                    return SystemIcons.Error;
                case MessageBoxIcon.Information:
                    return SystemIcons.Information;
                case MessageBoxIcon.Question:
                    return SystemIcons.Question;
                default:
                    return SystemIcons.Warning;
            }
        }

        internal sealed class DialogButtonSpec
        {
            public DialogButtonSpec(string text, DialogResult result, bool isAccept = false, bool isCancel = false)
            {
                Text = text ?? string.Empty;
                Result = result;
                IsAccept = isAccept;
                IsCancel = isCancel;
            }

            public string Text { get; }

            public DialogResult Result { get; }

            public bool IsAccept { get; }

            public bool IsCancel { get; }
        }
    }
}
