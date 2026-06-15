using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace OfficeAgent.ExcelAddIn.Dialogs
{
    internal sealed class BatchUploadProgressDialog : Form
    {
        private const int InitialDialogWidth = 960;
        private const int InitialDialogHeight = 700;
        private const int MinimumDialogWidth = 720;
        private const int MinimumDialogHeight = 460;
        private const int OuterPadding = 28;
        private const int StepMarkerColumnWidth = 48;
        private const int StepContentGap = 8;
        private const int StepSpacing = 28;
        private const int DetailsMaxHeight = 172;
        private const int DetailsMinHeight = 92;

        private readonly Label titleLabel;
        private readonly Button closeButton;
        private readonly Panel contentPanel;
        private readonly FlowLayoutPanel stepsPanel;
        private readonly Panel footerPanel;
        private readonly Panel headerPanel;
        private readonly List<StepRow> stepRows = new List<StepRow>();

        public BatchUploadProgressDialog(IEnumerable<BatchUploadProgressStep> steps)
        {
            if (steps == null)
            {
                throw new ArgumentNullException("steps");
            }

            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Dpi;
            Text = "批量上传";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(InitialDialogWidth, InitialDialogHeight);
            MinimumSize = new Size(MinimumDialogWidth, MinimumDialogHeight);

            titleLabel = new Label
            {
                Name = "titleLabel",
                AutoSize = false,
                Text = "批量上传",
                Margin = Padding.Empty,
                Font = new Font(Font.FontFamily, Font.Size + 3f, FontStyle.Regular),
            };

            headerPanel = new Panel
            {
                Name = "headerPanel",
                Dock = DockStyle.Top,
                Height = ScaleVertical(58),
                Padding = new Padding(OuterPadding, ScaleVertical(20), OuterPadding, 0),
            };
            headerPanel.Controls.Add(titleLabel);

            footerPanel = new Panel
            {
                Name = "footerPanel",
                Dock = DockStyle.Bottom,
                Height = ScaleVertical(76),
                Padding = new Padding(OuterPadding, ScaleVertical(10), OuterPadding, ScaleVertical(20)),
            };

            closeButton = new Button
            {
                Name = "closeButton",
                Text = "关闭",
                Anchor = AnchorStyles.Right | AnchorStyles.Bottom,
                MinimumSize = new Size(96, 34),
            };
            closeButton.Click += (sender, args) => Close();
            footerPanel.Controls.Add(closeButton);

            contentPanel = new Panel
            {
                Name = "contentPanel",
                Dock = DockStyle.Fill,
                AutoScroll = true,
                Padding = new Padding(OuterPadding, 0, OuterPadding, 0),
            };

            stepsPanel = new FlowLayoutPanel
            {
                Name = "stepsPanel",
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Top,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                Margin = Padding.Empty,
                Padding = Padding.Empty,
            };
            contentPanel.Controls.Add(stepsPanel);

            var stepList = steps.ToList();
            for (var index = 0; index < stepList.Count; index++)
            {
                var row = new StepRow(index + 1, stepList[index]);
                stepRows.Add(row);
                stepsPanel.Controls.Add(row);
            }

            Controls.Add(contentPanel);
            Controls.Add(footerPanel);
            Controls.Add(headerPanel);

            Layout += (sender, args) => UpdateResponsiveLayout();
            FontChanged += (sender, args) => UpdateResponsiveLayout();
        }

        public static BatchUploadProgressDialog CreateSample()
        {
            return new BatchUploadProgressDialog(new[]
            {
                new BatchUploadProgressStep(
                    "数据准备",
                    "加载字段映射和验证数据",
                    BatchUploadStepState.Completed),
                new BatchUploadProgressStep(
                    "字段验证",
                    "验证所有者、每日工作及其他字段",
                    BatchUploadStepState.Completed),
                new BatchUploadProgressStep(
                    "变更预览",
                    "已确认",
                    BatchUploadStepState.Completed,
                    "部分上传将上传48个单元格，跳过4个单元格。\r\n变更内容：\r\n0331test / taskFlowNode_13882098334 -> 测试\r\n0331test / taskFlowNode_13892195334 -> 1111\r\n0331test / taskFlowNode_13883074334 -> 东努沙登加拉\r\n0331test1 / SITEOWNER -> 15012344321\r\n0331test1 / taskFlowNode_13882098334 -> 1111\r\n0331test1 / taskFlowNode_13892195334 -> 1111"),
                new BatchUploadProgressStep(
                    "数据上传",
                    "正在上传至服务器",
                    BatchUploadStepState.Completed),
                new BatchUploadProgressStep(
                    "上传结果",
                    "上传完成",
                    BatchUploadStepState.Active,
                    "成功：48项变更\r\n上传完成。\r\n已提交单元格：48\r\n分块上传结果：\r\n分块1：成功（48个单元格）"),
            });
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            UpdateResponsiveLayout();
        }

        private void UpdateResponsiveLayout()
        {
            if (IsDisposed)
            {
                return;
            }

            var titleHeight = Math.Max(titleLabel.Font.Height + 8, ScaleVertical(30));
            var headerTopPadding = ScaleVertical(18);
            var headerBottomPadding = ScaleVertical(12);
            headerPanel.Height = headerTopPadding + titleHeight + headerBottomPadding;
            titleLabel.SetBounds(
                OuterPadding,
                headerTopPadding,
                Math.Max(120, ClientSize.Width - (OuterPadding * 2)),
                titleHeight);

            var closeButtonWidth = Math.Max(110, TextRenderer.MeasureText(closeButton.Text, closeButton.Font).Width + 56);
            var closeButtonHeight = Math.Max(36, TextRenderer.MeasureText(closeButton.Text, closeButton.Font).Height + 18);
            closeButton.Size = new Size(closeButtonWidth, closeButtonHeight);
            closeButton.Location = new Point(
                Math.Max(OuterPadding, footerPanel.ClientSize.Width - footerPanel.Padding.Right - closeButton.Width),
                Math.Max(footerPanel.Padding.Top, footerPanel.ClientSize.Height - footerPanel.Padding.Bottom - closeButton.Height));

            var scrollBarAllowance = contentPanel.VerticalScroll.Visible ? SystemInformation.VerticalScrollBarWidth : 0;
            var availableWidth = Math.Max(
                260,
                contentPanel.ClientSize.Width - contentPanel.Padding.Left - contentPanel.Padding.Right - scrollBarAllowance);
            stepsPanel.Width = availableWidth;

            foreach (var row in stepRows)
            {
                row.SetAvailableWidth(availableWidth);
            }
        }

        private int ScaleVertical(int value)
        {
            return Math.Max(value, (int)Math.Round(value * Font.Height / 15.0));
        }

        internal sealed class BatchUploadProgressStep
        {
            private readonly string title;
            private readonly string description;
            private readonly BatchUploadStepState state;
            private readonly string details;

            public BatchUploadProgressStep(string title, string description, BatchUploadStepState state, string details = null)
            {
                this.title = title ?? string.Empty;
                this.description = description ?? string.Empty;
                this.state = state;
                this.details = details ?? string.Empty;
            }

            public string Title
            {
                get { return title; }
            }

            public string Description
            {
                get { return description; }
            }

            public BatchUploadStepState State
            {
                get { return state; }
            }

            public string Details
            {
                get { return details; }
            }
        }

        internal enum BatchUploadStepState
        {
            Pending,
            Active,
            Completed,
            Warning,
            Error,
        }

        private sealed class StepRow : Panel
        {
            private readonly StepMarker marker;
            private readonly Panel linePanel;
            private readonly Label titleLabel;
            private readonly Label descriptionLabel;
            private readonly TextBox detailsTextBox;

            public StepRow(int stepNumber, BatchUploadProgressStep step)
            {
                if (step == null)
                {
                    throw new ArgumentNullException("step");
                }

                Name = "stepRow" + stepNumber;
                AutoSize = true;
                AutoSizeMode = AutoSizeMode.GrowAndShrink;
                Margin = new Padding(0, 0, 0, StepSpacing);
                Padding = Padding.Empty;

                marker = new StepMarker(stepNumber, step.State)
                {
                    Name = "stepMarker" + stepNumber,
                    Size = new Size(34, 34),
                    Location = Point.Empty,
                };

                linePanel = new Panel
                {
                    Name = "stepLine" + stepNumber,
                    BackColor = ResolveLineColor(step.State),
                    Width = 2,
                    Location = new Point(16, 40),
                    MinimumSize = new Size(2, 34),
                };

                titleLabel = new Label
                {
                    Name = "stepTitleLabel" + stepNumber,
                    AutoSize = false,
                    Text = step.Title,
                    Font = new Font(SystemFonts.MessageBoxFont.FontFamily, SystemFonts.MessageBoxFont.Size + 1.5f, FontStyle.Regular),
                    ForeColor = ResolveTitleColor(step.State),
                };

                descriptionLabel = new Label
                {
                    Name = "stepDescriptionLabel" + stepNumber,
                    AutoSize = false,
                    Text = step.Description,
                    ForeColor = Color.FromArgb(96, 96, 96),
                };

                if (!string.IsNullOrWhiteSpace(step.Details))
                {
                    detailsTextBox = new TextBox
                    {
                        Name = "stepDetailsTextBox" + stepNumber,
                        BorderStyle = BorderStyle.FixedSingle,
                        Multiline = true,
                        ReadOnly = true,
                        ScrollBars = ScrollBars.Vertical,
                        Text = step.Details,
                        WordWrap = false,
                    };
                    Controls.Add(detailsTextBox);
                }

                Controls.Add(marker);
                Controls.Add(linePanel);
                Controls.Add(titleLabel);
                Controls.Add(descriptionLabel);
            }

            public void SetAvailableWidth(int width)
            {
                Width = Math.Max(260, width);

                var contentWidth = Math.Max(160, Width - StepMarkerColumnWidth - StepContentGap);
                var contentLeft = StepMarkerColumnWidth + StepContentGap;
                var currentTop = 0;

                var titleHeight = MeasureWrappedHeight(titleLabel.Text, titleLabel.Font, contentWidth);
                titleLabel.SetBounds(contentLeft, currentTop, contentWidth, titleHeight);
                currentTop += titleHeight + 4;

                var descriptionHeight = MeasureWrappedHeight(descriptionLabel.Text, descriptionLabel.Font, contentWidth);
                descriptionLabel.SetBounds(contentLeft, currentTop, contentWidth, descriptionHeight);
                currentTop += descriptionHeight;

                if (detailsTextBox != null)
                {
                    currentTop += 8;
                    var detailsHeight = ResolveDetailsHeight(detailsTextBox.Text, detailsTextBox.Font, contentWidth);
                    detailsTextBox.SetBounds(contentLeft, currentTop, contentWidth, detailsHeight);
                    currentTop += detailsHeight;
                }

                Height = Math.Max(marker.Height + 48, currentTop);
                linePanel.Height = Math.Max(34, Height - marker.Height - 8);
            }

            public override Size GetPreferredSize(Size proposedSize)
            {
                return new Size(Width, Height);
            }

            private static int MeasureWrappedHeight(string text, Font font, int width)
            {
                var measured = TextRenderer.MeasureText(
                    text ?? string.Empty,
                    font,
                    new Size(Math.Max(1, width), int.MaxValue),
                    TextFormatFlags.WordBreak | TextFormatFlags.TextBoxControl);
                return Math.Max(font.Height, measured.Height);
            }

            private static int ResolveDetailsHeight(string text, Font font, int width)
            {
                var lineCount = string.IsNullOrEmpty(text)
                    ? 1
                    : text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None).Length;
                var desired = (font.Height * Math.Min(7, Math.Max(3, lineCount))) + 18;
                return Math.Max(DetailsMinHeight, Math.Min(DetailsMaxHeight, desired));
            }

            private static Color ResolveTitleColor(BatchUploadStepState state)
            {
                return state == BatchUploadStepState.Pending
                    ? Color.FromArgb(128, 128, 128)
                    : Color.FromArgb(32, 32, 32);
            }

            private static Color ResolveLineColor(BatchUploadStepState state)
            {
                return state == BatchUploadStepState.Completed || state == BatchUploadStepState.Active
                    ? Color.FromArgb(117, 181, 224)
                    : Color.FromArgb(232, 232, 232);
            }
        }

        private sealed class StepMarker : Control
        {
            private readonly int stepNumber;
            private readonly BatchUploadStepState state;

            public StepMarker(int stepNumber, BatchUploadStepState state)
            {
                this.stepNumber = stepNumber;
                this.state = state;
                DoubleBuffered = true;
                MinimumSize = new Size(30, 30);
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);

                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                var bounds = new Rectangle(1, 1, Width - 3, Height - 3);
                var fill = ResolveFillColor(state);
                var stroke = ResolveStrokeColor(state);

                using (var brush = new SolidBrush(fill))
                using (var pen = new Pen(stroke, 1.5f))
                {
                    e.Graphics.FillEllipse(brush, bounds);
                    e.Graphics.DrawEllipse(pen, bounds);
                }

                var text = ResolveText(state, stepNumber);
                var textColor = state == BatchUploadStepState.Completed || state == BatchUploadStepState.Active || state == BatchUploadStepState.Error
                    ? Color.White
                    : Color.FromArgb(96, 96, 96);
                TextRenderer.DrawText(
                    e.Graphics,
                    text,
                    Font,
                    bounds,
                    textColor,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine);
            }

            private static string ResolveText(BatchUploadStepState state, int stepNumber)
            {
                if (state == BatchUploadStepState.Completed)
                {
                    return "✓";
                }

                return stepNumber.ToString();
            }

            private static Color ResolveFillColor(BatchUploadStepState state)
            {
                switch (state)
                {
                    case BatchUploadStepState.Active:
                    case BatchUploadStepState.Completed:
                        return Color.FromArgb(0, 120, 215);
                    case BatchUploadStepState.Warning:
                        return Color.White;
                    case BatchUploadStepState.Error:
                        return Color.FromArgb(213, 48, 75);
                    default:
                        return Color.FromArgb(239, 239, 239);
                }
            }

            private static Color ResolveStrokeColor(BatchUploadStepState state)
            {
                switch (state)
                {
                    case BatchUploadStepState.Active:
                    case BatchUploadStepState.Completed:
                        return Color.FromArgb(0, 120, 215);
                    case BatchUploadStepState.Warning:
                    case BatchUploadStepState.Error:
                        return Color.FromArgb(213, 48, 75);
                    default:
                        return Color.FromArgb(204, 204, 204);
                }
            }
        }
    }
}
