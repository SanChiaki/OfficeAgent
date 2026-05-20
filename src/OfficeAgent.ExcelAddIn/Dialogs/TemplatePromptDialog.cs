using System;
using System.Drawing;
using System.Windows.Forms;

namespace OfficeAgent.ExcelAddIn.Dialogs
{
    internal sealed class TemplatePromptDialog : Form
    {
        private const int DefaultDialogWidth = 540;
        private const int DefaultDialogHeight = 220;
        private const int DialogPadding = 16;
        private const int IconColumnWidth = 48;

        private TemplatePromptDialog(
            string title,
            string message,
            MessageBoxIcon? icon,
            PromptOptions options,
            params DialogButtonSpec[] buttons)
        {
            if (buttons == null || buttons.Length == 0)
            {
                throw new ArgumentException("At least one button is required.", nameof(buttons));
            }

            options = options ?? new PromptOptions();

            Text = title ?? "ISDP";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(options.Width, options.Height);

            var root = new TableLayoutPanel
            {
                ColumnCount = 1,
                Dock = DockStyle.Fill,
                Padding = new Padding(DialogPadding),
                RowCount = 2,
            };
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle());

            var hasIcon = icon.HasValue;
            var content = new TableLayoutPanel
            {
                ColumnCount = hasIcon ? 2 : 1,
                Dock = DockStyle.Fill,
                RowCount = 1,
            };
            if (hasIcon)
            {
                content.ColumnStyles.Add(new ColumnStyle());
            }

            content.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

            if (hasIcon)
            {
                var iconBox = new PictureBox
                {
                    Image = ResolveIcon(icon.Value).ToBitmap(),
                    Margin = new Padding(0, 4, 16, 0),
                    Size = new Size(32, 32),
                    SizeMode = PictureBoxSizeMode.StretchImage,
                };
                content.Controls.Add(iconBox, 0, 0);
            }

            var messageWidth = Math.Max(120, options.Width - (DialogPadding * 2) - (hasIcon ? IconColumnWidth : 0));

            var messageLabel = new Label
            {
                AutoSize = true,
                Dock = DockStyle.Fill,
                Margin = Padding.Empty,
                MaximumSize = new Size(messageWidth, 0),
                Text = message ?? string.Empty,
            };

            if (options.EnableMessageScroll)
            {
                messageLabel.Dock = DockStyle.Top;
                var messagePanel = new Panel
                {
                    AutoScroll = true,
                    Dock = DockStyle.Fill,
                    Margin = Padding.Empty,
                };
                messagePanel.Controls.Add(messageLabel);
                content.Controls.Add(messagePanel, hasIcon ? 1 : 0, 0);
            }
            else
            {
                content.Controls.Add(messageLabel, hasIcon ? 1 : 0, 0);
            }

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
                    DialogResult = DialogResult.None,
                    Height = 30,
                    MinimumSize = new Size(88, 30),
                    Text = spec.Text ?? string.Empty,
                    Width = Math.Max(88, TextRenderer.MeasureText(spec.Text ?? string.Empty, SystemFonts.MessageBoxFont).Width + 24),
                };
                button.Click += (sender, e) =>
                {
                    spec.Invoke(this);
                    if (spec.Result != DialogResult.None)
                    {
                        DialogResult = spec.Result;
                        Close();
                    }
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
            return ShowPrompt(null, title, message, icon, null, buttons);
        }

        public static DialogResult ShowPrompt(
            IWin32Window owner,
            string title,
            string message,
            MessageBoxIcon? icon,
            PromptOptions options,
            params DialogButtonSpec[] buttons)
        {
            using (var dialog = new TemplatePromptDialog(title, message, icon, options, buttons))
            {
                return owner == null ? dialog.ShowDialog() : dialog.ShowDialog(owner);
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
            private readonly Action<IWin32Window> action;

            public DialogButtonSpec(
                string text,
                DialogResult result,
                bool isAccept = false,
                bool isCancel = false,
                Action<IWin32Window> action = null)
            {
                Text = text ?? string.Empty;
                Result = result;
                IsAccept = isAccept;
                IsCancel = isCancel;
                this.action = action;
            }

            public string Text { get; }

            public DialogResult Result { get; }

            public bool IsAccept { get; }

            public bool IsCancel { get; }

            public void Invoke(IWin32Window owner)
            {
                action?.Invoke(owner);
            }
        }

        internal sealed class PromptOptions
        {
            public int Width { get; set; } = DefaultDialogWidth;

            public int Height { get; set; } = DefaultDialogHeight;

            public bool EnableMessageScroll { get; set; }
        }
    }
}
