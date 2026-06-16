using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using OfficeAgent.ExcelAddIn.Localization;

namespace OfficeAgent.ExcelAddIn.Dialogs
{
    internal sealed class BatchUploadProgressDialog : Form
    {
        private const int InitialDialogWidth = 840;
        private const int InitialDialogHeight = 680;
        private const int MinimumDialogWidth = 680;
        private const int MinimumDialogHeight = 460;
        private const int OuterPadding = 28;
        private const int StepMarkerColumnWidth = 48;
        private const int StepContentGap = 8;
        private const int StepProgressRingSize = 46;
        private const int StepProgressRingGap = 24;
        private const int StepSpacing = 24;
        private const int ActiveProgressRingTopPadding = 14;
        private const int PreviewDetailsHeightMultiplier = 3;
        private const int PreviewDetailsHeightDivisor = 2;
        private const int DetailsMaxHeight = 240;
        private const int PreviewDetailsMaxHeight = 360;
        private const int DetailsMinHeight = 142;
        private const int FooterButtonGap = 12;

        private readonly Label titleLabel;
        private readonly Button uploadButton;
        private readonly Button cancelUploadButton;
        private readonly Button confirmButton;
        private readonly Panel contentPanel;
        private readonly FlowLayoutPanel stepsPanel;
        private readonly Panel footerPanel;
        private readonly Panel headerPanel;
        private readonly List<StepRow> stepRows = new List<StepRow>();
        private readonly ILocalizedDialogText strings;
        private readonly Stopwatch step2ToStep3Stopwatch = new Stopwatch();
        private bool previewHasUploadableContent = true;
        private long? step2ToStep3ElapsedMilliseconds;

        public event EventHandler UploadRequested;

        public event EventHandler UploadCanceled;

        public event EventHandler Confirmed;

        public long? Step2ToStep3ElapsedMilliseconds
        {
            get { return step2ToStep3ElapsedMilliseconds; }
        }

        public BatchUploadProgressDialog(IEnumerable<BatchUploadProgressStep> steps)
            : this(steps, (ILocalizedDialogText)null)
        {
        }

        public BatchUploadProgressDialog(IEnumerable<BatchUploadProgressStep> steps, HostLocalizedStrings strings)
            : this(steps, strings == null ? null : new HostLocalizedDialogText(strings))
        {
        }

        private BatchUploadProgressDialog(IEnumerable<BatchUploadProgressStep> steps, ILocalizedDialogText strings)
        {
            if (steps == null)
            {
                throw new ArgumentNullException("steps");
            }

            this.strings = strings ?? ResolveDefaultStrings();

            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Dpi;
            Text = this.strings.BatchUploadDialogTitle;
            StartPosition = FormStartPosition.CenterScreen;
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
                Text = this.strings.BatchUploadDialogTitle,
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

            uploadButton = new Button
            {
                Name = "uploadButton",
                Text = this.strings.BatchUploadUploadButtonText,
                Anchor = AnchorStyles.Right | AnchorStyles.Bottom,
                MinimumSize = new Size(96, 34),
            };
            uploadButton.Click += (sender, args) => OnUploadRequested();

            cancelUploadButton = new Button
            {
                Name = "cancelUploadButton",
                Text = this.strings.BatchUploadCancelButtonText,
                Anchor = AnchorStyles.Right | AnchorStyles.Bottom,
                MinimumSize = new Size(96, 34),
            };
            cancelUploadButton.Click += (sender, args) => OnUploadCanceled();

            confirmButton = new Button
            {
                Name = "confirmButton",
                Text = this.strings.BatchUploadConfirmButtonText,
                Anchor = AnchorStyles.Right | AnchorStyles.Bottom,
                MinimumSize = new Size(96, 34),
            };
            confirmButton.Click += (sender, args) => OnConfirmed();

            footerPanel.Controls.Add(uploadButton);
            footerPanel.Controls.Add(cancelUploadButton);
            footerPanel.Controls.Add(confirmButton);

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
            RefreshFooterButtons();

            Controls.Add(contentPanel);
            Controls.Add(footerPanel);
            Controls.Add(headerPanel);

            Layout += (sender, args) => UpdateResponsiveLayout();
            FontChanged += (sender, args) => UpdateResponsiveLayout();
        }

        public static BatchUploadProgressDialog CreateSample()
        {
            return CreateSample(HostLocalizedStrings.ForLocale("zh"));
        }

        public static BatchUploadProgressDialog CreateSample(string locale)
        {
            return CreateSample(new HostLocalizedDialogText(HostLocalizedStrings.ForLocale(locale)));
        }

        public static BatchUploadProgressDialog CreateSample(HostLocalizedStrings strings)
        {
            return CreateSample((ILocalizedDialogText)new HostLocalizedDialogText(strings));
        }

        private static BatchUploadProgressDialog CreateSample(ILocalizedDialogText strings)
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
            }, strings);
        }

        private static ILocalizedDialogText ResolveDefaultStrings()
        {
            return new HostLocalizedDialogText(Globals.ThisAddIn?.HostLocalizedStrings ?? HostLocalizedStrings.ForLocale("en"));
        }

        public void SetStepPending(int stepNumber, string title, string description, string details = null)
        {
            UpdateStep(stepNumber, title, description, BatchUploadStepState.Pending, details);
        }

        public void SetStepActive(int stepNumber, string title, string description, string details = null)
        {
            UpdateStep(stepNumber, title, description, BatchUploadStepState.Active, details);
        }

        public void SetStepCompleted(int stepNumber, string title, string description, string details = null)
        {
            UpdateStep(stepNumber, title, description, BatchUploadStepState.Completed, details);
        }

        public void SetStepWarning(int stepNumber, string title, string description, string details = null)
        {
            UpdateStep(stepNumber, title, description, BatchUploadStepState.Warning, details);
        }

        public void SetStepError(int stepNumber, string title, string description, string details = null)
        {
            UpdateStep(stepNumber, title, description, BatchUploadStepState.Error, details);
        }

        public void SetPreviewUploadAvailability(bool hasUploadableContent)
        {
            RunOnUiThread(() =>
            {
                previewHasUploadableContent = hasUploadableContent;
                RefreshFooterButtons();
                UpdateResponsiveLayout();
            });
        }

        public void UpdateStep(int stepNumber, string title, string description, BatchUploadStepState state, string details = null)
        {
            RunOnUiThread(() =>
            {
                CaptureStep2ToStep3ElapsedTime(stepNumber, state);

                if (state == BatchUploadStepState.Active)
                {
                    ClearActiveStepsExcept(stepNumber);
                }

                var row = ResolveStepRow(stepNumber);
                row.UpdateStep(new BatchUploadProgressStep(title, description, state, details));
                RefreshFooterButtons();
                UpdateResponsiveLayout();
            });
        }

        private void CaptureStep2ToStep3ElapsedTime(int stepNumber, BatchUploadStepState state)
        {
            if (state != BatchUploadStepState.Active)
            {
                return;
            }

            if (stepNumber == 2)
            {
                step2ToStep3ElapsedMilliseconds = null;
                step2ToStep3Stopwatch.Restart();
                return;
            }

            if (stepNumber == 3 && step2ToStep3Stopwatch.IsRunning)
            {
                step2ToStep3Stopwatch.Stop();
                step2ToStep3ElapsedMilliseconds = step2ToStep3Stopwatch.ElapsedMilliseconds;
            }
        }

        public void AppendStepDetails(int stepNumber, string details)
        {
            if (string.IsNullOrEmpty(details))
            {
                return;
            }

            RunOnUiThread(() =>
            {
                var row = ResolveStepRow(stepNumber);
                row.AppendDetails(details);
                UpdateResponsiveLayout();
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

            var uploadButtonWidth = Math.Max(110, TextRenderer.MeasureText(uploadButton.Text, uploadButton.Font).Width + 56);
            var cancelButtonWidth = Math.Max(110, TextRenderer.MeasureText(cancelUploadButton.Text, cancelUploadButton.Font).Width + 56);
            var confirmButtonWidth = Math.Max(110, TextRenderer.MeasureText(confirmButton.Text, confirmButton.Font).Width + 56);
            var footerButtonHeight = Math.Max(
                36,
                Math.Max(
                    Math.Max(
                        TextRenderer.MeasureText(uploadButton.Text, uploadButton.Font).Height,
                        TextRenderer.MeasureText(cancelUploadButton.Text, cancelUploadButton.Font).Height),
                    TextRenderer.MeasureText(confirmButton.Text, confirmButton.Font).Height) + 18);
            uploadButton.Size = new Size(uploadButtonWidth, footerButtonHeight);
            cancelUploadButton.Size = new Size(cancelButtonWidth, footerButtonHeight);
            confirmButton.Size = new Size(confirmButtonWidth, footerButtonHeight);
            LayoutFooterButtons(footerButtonHeight);

            var hadVerticalScroll = contentPanel.VerticalScroll.Visible;
            LayoutStepRows();

            if (contentPanel.VerticalScroll.Visible != hadVerticalScroll)
            {
                LayoutStepRows();
            }
        }

        private void LayoutStepRows()
        {
            var availableWidth = Math.Max(
                260,
                contentPanel.ClientSize.Width - contentPanel.Padding.Left - contentPanel.Padding.Right);
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

        private void OnUploadRequested()
        {
            var handler = UploadRequested;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }

        private void OnUploadCanceled()
        {
            var handler = UploadCanceled;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }

            Close();
        }

        private void OnConfirmed()
        {
            var handler = Confirmed;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }

            Close();
        }

        private void RefreshFooterButtons()
        {
            var activeStep = ResolveActiveStepNumber();
            uploadButton.Visible = activeStep == 3 && previewHasUploadableContent;
            cancelUploadButton.Visible = activeStep == 1 || activeStep == 2 || (activeStep == 3 && previewHasUploadableContent) || activeStep == 4;
            confirmButton.Visible = (activeStep == 3 && !previewHasUploadableContent) || ((activeStep == 0 || activeStep == 5) && IsResultStepReadyForConfirmation());
            LayoutFooterButtons(Math.Max(uploadButton.Height, Math.Max(cancelUploadButton.Height, confirmButton.Height)));
        }

        private bool IsResultStepReadyForConfirmation()
        {
            if (stepRows.Count < 5)
            {
                return false;
            }

            return stepRows[4].State != BatchUploadStepState.Pending;
        }

        private int ResolveActiveStepNumber()
        {
            for (var index = 0; index < stepRows.Count; index++)
            {
                if (stepRows[index].State == BatchUploadStepState.Active)
                {
                    return index + 1;
                }
            }

            return 0;
        }

        private void ClearActiveStepsExcept(int activeStepNumber)
        {
            for (var index = 0; index < stepRows.Count; index++)
            {
                var stepNumber = index + 1;
                var row = stepRows[index];
                if (stepNumber != activeStepNumber && row.State == BatchUploadStepState.Active)
                {
                    row.UpdateStep(new BatchUploadProgressStep(
                        row.Title,
                        row.Description,
                        BatchUploadStepState.Completed,
                        row.Details));
                }
            }
        }

        private void LayoutFooterButtons(int buttonHeight)
        {
            var visibleButtons = new List<Button>();
            if (uploadButton.Visible)
            {
                visibleButtons.Add(uploadButton);
            }

            if (cancelUploadButton.Visible)
            {
                visibleButtons.Add(cancelUploadButton);
            }

            if (confirmButton.Visible)
            {
                visibleButtons.Add(confirmButton);
            }

            var top = Math.Max(
                footerPanel.Padding.Top,
                footerPanel.ClientSize.Height - footerPanel.Padding.Bottom - buttonHeight);
            var right = footerPanel.ClientSize.Width - footerPanel.Padding.Right;
            for (var index = visibleButtons.Count - 1; index >= 0; index--)
            {
                var button = visibleButtons[index];
                button.Location = new Point(Math.Max(OuterPadding, right - button.Width), top);
                right = button.Left - FooterButtonGap;
            }
        }

        private StepRow ResolveStepRow(int stepNumber)
        {
            if (stepNumber < 1 || stepNumber > stepRows.Count)
            {
                throw new ArgumentOutOfRangeException("stepNumber");
            }

            return stepRows[stepNumber - 1];
        }

        private void RunOnUiThread(Action action)
        {
            if (action == null)
            {
                throw new ArgumentNullException("action");
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
            private StepProgressRing progressRing;
            private TextBox detailsTextBox;
            private BatchUploadProgressStep step;
            private readonly int stepNumber;

            public StepRow(int stepNumber, BatchUploadProgressStep step)
            {
                if (step == null)
                {
                    throw new ArgumentNullException("step");
                }

                this.step = step;
                this.stepNumber = stepNumber;
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

                Controls.Add(marker);
                Controls.Add(linePanel);
                Controls.Add(titleLabel);
                Controls.Add(descriptionLabel);

                RefreshDetailsTextBox();
                RefreshProgressRing();
            }

            public void UpdateStep(BatchUploadProgressStep nextStep)
            {
                if (nextStep == null)
                {
                    throw new ArgumentNullException("nextStep");
                }

                step = nextStep;
                marker.State = nextStep.State;
                linePanel.BackColor = ResolveLineColor(nextStep.State);
                titleLabel.Text = nextStep.Title;
                titleLabel.ForeColor = ResolveTitleColor(nextStep.State);
                descriptionLabel.Text = nextStep.Description;
                RefreshDetailsTextBox();
                RefreshProgressRing();
                Invalidate(true);
            }

            public void AppendDetails(string details)
            {
                var mergedDetails = string.IsNullOrEmpty(step.Details)
                    ? details
                    : step.Details + "\r\n" + details;
                UpdateStep(new BatchUploadProgressStep(step.Title, step.Description, step.State, mergedDetails));
            }

            public BatchUploadStepState State
            {
                get { return step.State; }
            }

            public string Title
            {
                get { return step.Title; }
            }

            public string Description
            {
                get { return step.Description; }
            }

            public string Details
            {
                get { return step.Details; }
            }

            public void SetAvailableWidth(int width)
            {
                Width = Math.Max(260, width);

                var hasProgressRing = progressRing != null;
                var ringLeft = hasProgressRing
                    ? Math.Max(
                        StepMarkerColumnWidth + StepContentGap + 160 + StepProgressRingGap,
                        Width - StepProgressRingSize)
                    : Width;
                var contentWidth = Math.Max(
                    160,
                    ringLeft - (hasProgressRing ? StepProgressRingGap : 0) - StepMarkerColumnWidth - StepContentGap);
                var contentLeft = StepMarkerColumnWidth + StepContentGap;
                var currentTop = hasProgressRing ? ActiveProgressRingTopPadding : 0;

                var titleHeight = MeasureWrappedHeight(titleLabel.Text, titleLabel.Font, contentWidth);
                titleLabel.SetBounds(contentLeft, currentTop, contentWidth, titleHeight);
                currentTop += titleHeight + 4;

                var descriptionHeight = MeasureWrappedHeight(descriptionLabel.Text, descriptionLabel.Font, contentWidth);
                descriptionLabel.SetBounds(contentLeft, currentTop, contentWidth, descriptionHeight);
                currentTop += descriptionHeight;

                if (detailsTextBox != null)
                {
                    currentTop += 8;
                    var detailsHeight = ResolveDetailsHeight(detailsTextBox.Text, detailsTextBox.Font, contentWidth, stepNumber);
                    detailsTextBox.SetBounds(contentLeft, currentTop, contentWidth, detailsHeight);
                    currentTop += detailsHeight;
                }

                Height = Math.Max(marker.Height + 48, currentTop);
                linePanel.Height = Math.Max(34, Height - marker.Height - 8);
                if (progressRing != null)
                {
                    var ringTop = Math.Max(0, titleLabel.Top + ((titleLabel.Height - progressRing.Height) / 2));
                    progressRing.Location = new Point(
                        Math.Max(contentLeft + contentWidth + StepProgressRingGap, Width - StepProgressRingSize),
                        ringTop);
                }
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

            private static int ResolveDetailsHeight(string text, Font font, int width, int stepNumber)
            {
                var lineCount = string.IsNullOrEmpty(text)
                    ? 1
                    : text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None).Length;
                var desired = (font.Height * Math.Min(7, Math.Max(3, lineCount))) + 18;
                if (stepNumber == 3)
                {
                    desired = desired * PreviewDetailsHeightMultiplier / PreviewDetailsHeightDivisor;
                    return Math.Max(DetailsMinHeight, Math.Min(PreviewDetailsMaxHeight, desired));
                }

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

            private void RefreshDetailsTextBox()
            {
                var hasDetails = (stepNumber == 3 || stepNumber == 5) && !string.IsNullOrWhiteSpace(step.Details);
                if (!hasDetails)
                {
                    if (detailsTextBox != null)
                    {
                        Controls.Remove(detailsTextBox);
                        detailsTextBox.Dispose();
                        detailsTextBox = null;
                    }

                    return;
                }

                if (detailsTextBox == null)
                {
                    detailsTextBox = new TextBox
                    {
                        Name = "stepDetailsTextBox" + stepNumber,
                        BackColor = Color.White,
                        BorderStyle = BorderStyle.FixedSingle,
                        Multiline = true,
                        ReadOnly = true,
                        ScrollBars = ScrollBars.Both,
                        WordWrap = false,
                    };
                    Controls.Add(detailsTextBox);
                    detailsTextBox.BringToFront();
                }

                detailsTextBox.Text = step.Details;
            }

            private void RefreshProgressRing()
            {
                if (step.State != BatchUploadStepState.Active)
                {
                    if (progressRing != null)
                    {
                        Controls.Remove(progressRing);
                        progressRing.Dispose();
                        progressRing = null;
                    }

                    return;
                }

                if (progressRing == null)
                {
                    progressRing = new StepProgressRing(step.State)
                    {
                        Name = "stepProgressRing" + stepNumber,
                        Size = new Size(StepProgressRingSize, StepProgressRingSize),
                    };
                    Controls.Add(progressRing);
                    progressRing.BringToFront();
                }
            }
        }

        private sealed class StepMarker : Control
        {
            private readonly int stepNumber;
            private BatchUploadStepState state;

            public StepMarker(int stepNumber, BatchUploadStepState state)
            {
                this.stepNumber = stepNumber;
                this.state = state;
                DoubleBuffered = true;
                MinimumSize = new Size(30, 30);
            }

            public BatchUploadStepState State
            {
                get { return state; }
                set
                {
                    if (state == value)
                    {
                        return;
                    }

                    state = value;
                    Invalidate();
                }
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

        private sealed class StepProgressRing : Control
        {
            private readonly BatchUploadStepState state;
            private readonly Timer animationTimer;
            private int startAngle = -90;

            public StepProgressRing(BatchUploadStepState state)
            {
                this.state = state;
                DoubleBuffered = true;
                MinimumSize = new Size(42, 42);
                animationTimer = new Timer
                {
                    Interval = 90,
                };
                animationTimer.Tick += (sender, args) => AdvanceAnimationFrame();
                animationTimer.Start();
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);

                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                var strokeWidth = Math.Max(2f, Width / 18f);
                var inset = (int)Math.Ceiling(strokeWidth) + 2;
                var bounds = new Rectangle(inset, inset, Width - (inset * 2) - 1, Height - (inset * 2) - 1);

                using (var dashedPen = new Pen(Color.FromArgb(197, 214, 232), strokeWidth))
                using (var progressPen = new Pen(ResolveProgressColor(state), strokeWidth))
                {
                    dashedPen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dot;
                    dashedPen.StartCap = System.Drawing.Drawing2D.LineCap.Round;
                    dashedPen.EndCap = System.Drawing.Drawing2D.LineCap.Round;
                    progressPen.StartCap = System.Drawing.Drawing2D.LineCap.Round;
                    progressPen.EndCap = System.Drawing.Drawing2D.LineCap.Round;

                    e.Graphics.DrawEllipse(dashedPen, bounds);
                    e.Graphics.DrawArc(progressPen, bounds, startAngle, ResolveSweepAngle(state));
                }
            }

            protected override void Dispose(bool disposing)
            {
                if (disposing)
                {
                    animationTimer.Stop();
                    animationTimer.Dispose();
                }

                base.Dispose(disposing);
            }

            private void AdvanceAnimationFrame()
            {
                startAngle = (startAngle + 16) % 360;
                Invalidate();
            }

            private static int ResolveSweepAngle(BatchUploadStepState state)
            {
                switch (state)
                {
                    case BatchUploadStepState.Active:
                        return 72;
                    default:
                        return 48;
                }
            }

            private static Color ResolveProgressColor(BatchUploadStepState state)
            {
                switch (state)
                {
                    case BatchUploadStepState.Active:
                        return Color.FromArgb(0, 120, 215);
                    default:
                        return Color.FromArgb(154, 169, 184);
                }
            }
        }

        private interface ILocalizedDialogText
        {
            string BatchUploadDialogTitle { get; }

            string BatchUploadUploadButtonText { get; }

            string BatchUploadCancelButtonText { get; }

            string BatchUploadConfirmButtonText { get; }
        }

        private sealed class HostLocalizedDialogText : ILocalizedDialogText
        {
            private readonly HostLocalizedStrings strings;

            public HostLocalizedDialogText(HostLocalizedStrings strings)
            {
                this.strings = strings ?? HostLocalizedStrings.ForLocale("en");
            }

            public string BatchUploadDialogTitle
            {
                get { return strings.BatchUploadDialogTitle; }
            }

            public string BatchUploadUploadButtonText
            {
                get { return strings.BatchUploadUploadButtonText; }
            }

            public string BatchUploadCancelButtonText
            {
                get { return strings.BatchUploadCancelButtonText; }
            }

            public string BatchUploadConfirmButtonText
            {
                get { return strings.BatchUploadConfirmButtonText; }
            }
        }

    }
}
