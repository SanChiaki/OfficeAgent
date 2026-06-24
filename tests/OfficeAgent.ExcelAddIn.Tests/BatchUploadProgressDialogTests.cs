using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class BatchUploadProgressDialogTests
    {
        [Fact]
        public void DialogKeepsScrollableContentAndFixedFooterWhenFontScalesUp()
        {
            RunInSta(() =>
            {
                using (var dialog = CreateDialog())
                using (var scaledFont = new Font(dialog.Font.FontFamily, dialog.Font.Size * 1.55f, dialog.Font.Style))
                {
                    ApplyFont(dialog, scaledFont);
                    dialog.CreateControl();
                    dialog.PerformLayout();

                    var contentPanel = FindControl<Panel>(dialog, "contentPanel");
                    var footerPanel = FindControl<Panel>(dialog, "footerPanel");
                    var stepsPanel = FindControl<FlowLayoutPanel>(dialog, "stepsPanel");
                    var resultTextBox = FindControl<TextBox>(dialog, "stepDetailsTextBox5");

                    Assert.DoesNotContain(EnumerateControls(dialog), control => string.Equals(control.Name, "closeGlyphButton", StringComparison.Ordinal));
                    Assert.DoesNotContain(EnumerateControls(dialog), control => string.Equals(control.Name, "closeButton", StringComparison.Ordinal));
                    Assert.True(contentPanel.AutoScroll);
                    Assert.Equal(DockStyle.Fill, contentPanel.Dock);
                    Assert.Equal(DockStyle.Bottom, footerPanel.Dock);
                    Assert.Equal(FlowDirection.TopDown, stepsPanel.FlowDirection);
                    Assert.True(stepsPanel.AutoSize);
                    Assert.True(resultTextBox.Multiline);
                    Assert.Equal(ScrollBars.Both, resultTextBox.ScrollBars);
                    Assert.True(resultTextBox.Width >= contentPanel.ClientSize.Width - 160);

                    AssertNoVisibleTextControlClips(dialog);
                }
            });
        }

        [Fact]
        public void DialogUsesCompactScreenCenteredLayoutWithLargerWhiteDetailsBox()
        {
            RunInSta(() =>
            {
                using (var dialog = CreateDialog())
                {
                    dialog.CreateControl();
                    dialog.PerformLayout();

                    var previewTextBox = FindControl<TextBox>(dialog, "stepDetailsTextBox3");
                    var resultTextBox = FindControl<TextBox>(dialog, "stepDetailsTextBox5");

                    Assert.Equal(FormStartPosition.CenterScreen, dialog.StartPosition);
                    Assert.True(dialog.ClientSize.Width <= 860, $"Dialog width should be compact, actual: {dialog.ClientSize.Width}.");
                    Assert.True(dialog.ClientSize.Height <= 680, $"Dialog height should stay compact, actual: {dialog.ClientSize.Height}.");
                    Assert.Equal(Color.White, previewTextBox.BackColor);
                    Assert.Equal(Color.White, resultTextBox.BackColor);
                    Assert.Equal(ScrollBars.Both, previewTextBox.ScrollBars);
                    Assert.Equal(ScrollBars.Both, resultTextBox.ScrollBars);
                    Assert.True(previewTextBox.Height >= 140, $"Preview details box should be larger, actual: {previewTextBox.Height}.");
                    Assert.True(resultTextBox.Height >= 140, $"Result details box should be larger, actual: {resultTextBox.Height}.");
                    Assert.True(previewTextBox.Height >= resultTextBox.Height * 1.45, $"Preview details box should be about 1.5x the result box. Preview={previewTextBox.Height}, Result={resultTextBox.Height}.");
                }
            });
        }

        [Fact]
        public void PreviewDetailsRelayoutsProgressRingWhenScrollbarAppears()
        {
            RunInSta(() =>
            {
                using (var dialog = CreateDialog())
                {
                    ShowOffscreen(dialog);
                    InvokeDialogMethod(dialog, "SetStepActive", 3, "变更预览", "等待生成预览", null);
                    Application.DoEvents();

                    InvokeDialogMethod(
                        dialog,
                        "SetStepActive",
                        3,
                        "变更预览",
                        "请确认本次将上传的内容",
                        BuildPreviewDetails());

                    var thirdRow = FindControl<Control>(dialog, "stepRow3");
                    var ring = FindControl<Control>(thirdRow, "stepProgressRing3");
                    var title = FindControl<Label>(thirdRow, "stepTitleLabel3");
                    var details = FindControl<TextBox>(thirdRow, "stepDetailsTextBox3");

                    Assert.True(ring.Right <= thirdRow.Width, $"Progress ring should stay inside row bounds. RingRight={ring.Right}, RowWidth={thirdRow.Width}.");
                    Assert.True(details.Right < ring.Left, $"Details box should not overlap the progress ring. DetailsRight={details.Right}, RingLeft={ring.Left}.");
                    Assert.True(Math.Abs(CenterY(ring) - CenterY(title)) <= 4, $"Progress ring should align with the step title. RingCenter={CenterY(ring)}, TitleCenter={CenterY(title)}.");
                }
            });
        }

        [Fact]
        public void DialogFooterUsesConfirmButtonForResultStep()
        {
            RunInSta(() =>
            {
                using (var dialog = CreateDialog())
                {
                    dialog.CreateControl();
                    dialog.PerformLayout();

                    var confirmButton = FindControl<Button>(dialog, "confirmButton");

                    Assert.Equal("确认", confirmButton.Text);
                    Assert.Null(TryFindControl<Button>(dialog, "uploadButton"));
                    Assert.Null(TryFindControl<Button>(dialog, "cancelUploadButton"));
                    Assert.DoesNotContain(EnumerateControls(dialog), control => string.Equals(control.Name, "closeButton", StringComparison.Ordinal));
                }
            });
        }

        [Fact]
        public void DialogUsesLocalizedChromeAndFooterButtons()
        {
            RunInSta(() =>
            {
                using (var dialog = CreateDialog("en"))
                {
                    dialog.CreateControl();
                    dialog.PerformLayout();

                    Assert.Equal("Batch upload", dialog.Text);
                    Assert.Equal("Batch upload", FindControl<Label>(dialog, "titleLabel").Text);
                    Assert.Equal("Upload", FindControl<Button>(dialog, "uploadButton").Text);
                    Assert.Equal("Cancel", FindControl<Button>(dialog, "cancelUploadButton").Text);
                    Assert.Equal("Confirm", FindControl<Button>(dialog, "confirmButton").Text);
                }
            });
        }

        [Fact]
        public void DialogFooterUsesConfirmButtonWhenResultStepIsFinished()
        {
            RunInSta(() =>
            {
                using (var dialog = CreateDialog())
                {
                    dialog.CreateControl();
                    dialog.PerformLayout();

                    InvokeDialogMethod(dialog, "SetStepCompleted", 5, "上传结果", "上传完成", "成功：48项变更");
                    var completedButtons = VisibleFooterButtons(dialog);
                    Assert.Single(completedButtons);
                    Assert.Equal("确认", completedButtons.Single().Text);

                    InvokeDialogMethod(dialog, "SetStepError", 5, "上传结果", "上传失败", "请查看日志确认失败原因");
                    var errorButtons = VisibleFooterButtons(dialog);
                    Assert.Single(errorButtons);
                    Assert.Equal("确认", errorButtons.Single().Text);
                }
            });
        }

        [Fact]
        public void DialogFooterButtonsChangeWithActiveStep()
        {
            RunInSta(() =>
            {
                using (var dialog = CreateDialog())
                {
                    dialog.CreateControl();
                    dialog.PerformLayout();

                    InvokeDialogMethod(dialog, "SetStepActive", 1, "数据准备", "正在读取 Excel 可见选区", null);
                    Assert.Single(VisibleFooterButtons(dialog));
                    Assert.Equal("取消", VisibleFooterButtons(dialog).Single().Text);
                    Assert.Null(TryFindControl<TextBox>(dialog, "stepDetailsTextBox1"));

                    InvokeDialogMethod(dialog, "SetStepActive", 2, "字段验证", "正在验证字段", "第 2 步详情不应该显示");
                    Assert.Single(VisibleFooterButtons(dialog));
                    Assert.Equal("取消", VisibleFooterButtons(dialog).Single().Text);
                    Assert.Null(TryFindControl<TextBox>(dialog, "stepDetailsTextBox2"));

                    InvokeDialogMethod(dialog, "SetStepActive", 3, "变更预览", "确认本次上传内容", "将上传 48 个单元格");
                    var previewButtons = VisibleFooterButtons(dialog);
                    Assert.Equal(new[] { "上传", "取消" }, previewButtons.Select(button => button.Text).ToArray());
                    Assert.DoesNotContain(previewButtons, button => string.Equals(button.Text, "确认", StringComparison.Ordinal));
                    Assert.NotNull(TryFindControl<TextBox>(dialog, "stepDetailsTextBox3"));

                    InvokeDialogMethod(dialog, "SetStepActive", 4, "数据上传", "正在上传至服务器", "这个详情不应该显示");
                    Assert.Single(VisibleFooterButtons(dialog));
                    Assert.Equal("取消", VisibleFooterButtons(dialog).Single().Text);
                    Assert.NotNull(TryFindControl<TextBox>(dialog, "stepDetailsTextBox3"));
                    Assert.Null(TryFindControl<TextBox>(dialog, "stepDetailsTextBox4"));

                    InvokeDialogMethod(dialog, "SetStepActive", 5, "上传结果", "上传完成", "成功：48项变更");
                    Assert.Single(VisibleFooterButtons(dialog));
                    Assert.Equal("确认", VisibleFooterButtons(dialog).Single().Text);
                    Assert.NotNull(TryFindControl<TextBox>(dialog, "stepDetailsTextBox3"));
                    Assert.NotNull(TryFindControl<TextBox>(dialog, "stepDetailsTextBox5"));
                }
            });
        }

        [Fact]
        public void PreviewStepShowsConfirmOnlyWhenNoUploadableContentExists()
        {
            RunInSta(() =>
            {
                using (var dialog = CreateDialog())
                {
                    InvokeDialogMethod(dialog, "SetPreviewUploadAvailability", false);
                    dialog.CreateControl();
                    dialog.PerformLayout();

                    InvokeDialogMethod(dialog, "SetStepActive", 3, "变更预览", "没有可上传内容", "所选内容均不满足上传条件");
                    var previewButtons = VisibleFooterButtons(dialog);

                    Assert.Single(previewButtons);
                    Assert.Equal("确认", previewButtons.Single().Text);
                    Assert.Null(TryFindControl<Button>(dialog, "uploadButton"));
                    Assert.Null(TryFindControl<Button>(dialog, "cancelUploadButton"));
                }
            });
        }

        [Fact]
        public void PreviewConfirmButtonClosesDialogWhenNoUploadableContentExists()
        {
            RunInSta(() =>
            {
                using (var dialog = CreateDialog())
                {
                    var closed = false;
                    dialog.FormClosed += (sender, args) => closed = true;
                    InvokeDialogMethod(dialog, "SetPreviewUploadAvailability", false);
                    ShowOffscreen(dialog);
                    InvokeDialogMethod(dialog, "SetStepActive", 3, "变更预览", "没有可上传内容", "所选内容均不满足上传条件");

                    FindControl<Button>(dialog, "confirmButton").PerformClick();
                    Application.DoEvents();

                    Assert.True(closed);
                }
            });
        }

        [Fact]
        public void DialogFooterButtonsRaiseActionEvents()
        {
            RunInSta(() =>
            {
                using (var dialog = CreateDialog())
                {
                    var uploadRequested = 0;
                    var uploadRequestedEvent = dialog.GetType().GetEvent("UploadRequested")
                        ?? throw new InvalidOperationException("UploadRequested event was not found.");
                    uploadRequestedEvent.AddEventHandler(dialog, new EventHandler((sender, args) => uploadRequested++));

                    dialog.CreateControl();
                    dialog.PerformLayout();

                    InvokeDialogMethod(dialog, "SetStepActive", 3, "变更预览", "确认本次上传内容", "将上传 48 个单元格");
                    FindControl<Button>(dialog, "uploadButton").PerformClick();

                    Assert.Equal(1, uploadRequested);
                }

                using (var dialog = CreateDialog())
                {
                    var uploadCanceled = 0;
                    var uploadCanceledEvent = dialog.GetType().GetEvent("UploadCanceled")
                        ?? throw new InvalidOperationException("UploadCanceled event was not found.");
                    uploadCanceledEvent.AddEventHandler(dialog, new EventHandler((sender, args) => uploadCanceled++));

                    dialog.CreateControl();
                    dialog.PerformLayout();
                    InvokeDialogMethod(dialog, "SetStepActive", 3, "变更预览", "确认本次上传内容", "将上传 48 个单元格");
                    FindControl<Button>(dialog, "cancelUploadButton").PerformClick();

                    Assert.Equal(1, uploadCanceled);
                }

                using (var dialog = CreateDialog())
                {
                    var confirmed = 0;
                    var confirmedEvent = dialog.GetType().GetEvent("Confirmed")
                        ?? throw new InvalidOperationException("Confirmed event was not found.");
                    confirmedEvent.AddEventHandler(dialog, new EventHandler((sender, args) => confirmed++));

                    dialog.CreateControl();
                    dialog.PerformLayout();
                    InvokeDialogMethod(dialog, "SetStepActive", 5, "上传结果", "上传完成", "成功：48项变更");
                    FindControl<Button>(dialog, "confirmButton").PerformClick();

                    Assert.Equal(1, confirmed);
                }
            });
        }

        [Fact]
        public void ConfirmButtonClosesDialog()
        {
            RunInSta(() =>
            {
                using (var dialog = CreateDialog())
                {
                    var closed = false;
                    dialog.FormClosed += (sender, args) => closed = true;
                    ShowOffscreen(dialog);

                    FindControl<Button>(dialog, "confirmButton").PerformClick();
                    Application.DoEvents();

                    Assert.True(closed);
                }
            });
        }

        [Fact]
        public void CancelButtonClosesDialogBeforeUploadStarts()
        {
            RunInSta(() =>
            {
                using (var dialog = CreateDialog())
                {
                    var closed = false;
                    dialog.FormClosed += (sender, args) => closed = true;
                    ShowOffscreen(dialog);
                    InvokeDialogMethod(dialog, "SetStepActive", 1, "数据准备", "正在读取 Excel 可见选区", null);

                    FindControl<Button>(dialog, "cancelUploadButton").PerformClick();
                    Application.DoEvents();

                    Assert.True(closed);
                }
            });
        }

        [Fact]
        public void CancelButtonClosesDialogWhileUploading()
        {
            RunInSta(() =>
            {
                using (var dialog = CreateDialog())
                {
                    var closed = false;
                    dialog.FormClosed += (sender, args) => closed = true;
                    ShowOffscreen(dialog);
                    InvokeDialogMethod(dialog, "SetStepActive", 4, "数据上传", "正在上传至服务器", null);

                    FindControl<Button>(dialog, "cancelUploadButton").PerformClick();
                    Application.DoEvents();

                    Assert.True(closed);
                }
            });
        }

        [Fact]
        public void DialogUsesContentDrivenStepRowsInsteadOfEqualHeightBuckets()
        {
            RunInSta(() =>
            {
                using (var dialog = CreateDialog())
                {
                    dialog.CreateControl();
                    dialog.PerformLayout();

                    var stepsPanel = FindControl<FlowLayoutPanel>(dialog, "stepsPanel");
                    var stepRows = stepsPanel.Controls.Cast<Control>().Where(control => control.Name.StartsWith("stepRow", StringComparison.Ordinal)).ToArray();

                    Assert.Equal(5, stepRows.Length);
                    Assert.All(stepRows, row => Assert.True(row.AutoSize));
                    Assert.Contains(stepRows, row => row.Height > stepRows.Min(candidate => candidate.Height));
                }
            });
        }

        [Fact]
        public void StepMarkersUseCheckOnlyForCompletedSteps()
        {
            var dialogType = LoadAddInAssembly().GetType(
                "OfficeAgent.ExcelAddIn.Dialogs.BatchUploadProgressDialog",
                throwOnError: true);
            var markerType = dialogType.GetNestedType("StepMarker", BindingFlags.NonPublic)
                ?? throw new InvalidOperationException("StepMarker was not found.");
            var stateType = dialogType.GetNestedType("BatchUploadStepState", BindingFlags.NonPublic)
                ?? throw new InvalidOperationException("BatchUploadStepState was not found.");
            var resolveText = markerType.GetMethod("ResolveText", BindingFlags.NonPublic | BindingFlags.Static)
                ?? throw new InvalidOperationException("StepMarker.ResolveText was not found.");

            Assert.Equal("✓", ResolveMarkerText(resolveText, stateType, "Completed", 1));
            Assert.Equal("2", ResolveMarkerText(resolveText, stateType, "Active", 2));
            Assert.Equal("3", ResolveMarkerText(resolveText, stateType, "Pending", 3));
            Assert.Equal("4", ResolveMarkerText(resolveText, stateType, "Warning", 4));
            Assert.Equal("5", ResolveMarkerText(resolveText, stateType, "Error", 5));
        }

        [Fact]
        public void StepRowsShowProgressRingOnlyForActiveStep()
        {
            RunInSta(() =>
            {
                using (var dialog = CreateDialog())
                {
                    dialog.CreateControl();
                    dialog.PerformLayout();

                    var stepsPanel = FindControl<FlowLayoutPanel>(dialog, "stepsPanel");
                    var stepRows = stepsPanel.Controls.Cast<Control>()
                        .Where(control => control.Name.StartsWith("stepRow", StringComparison.Ordinal))
                        .ToArray();

                    Assert.Equal(5, stepRows.Length);
                    Assert.Null(TryFindControl<Control>(stepRows[0], "stepProgressRing1"));
                    Assert.Null(TryFindControl<Control>(stepRows[1], "stepProgressRing2"));
                    Assert.Null(TryFindControl<Control>(stepRows[2], "stepProgressRing3"));
                    Assert.Null(TryFindControl<Control>(stepRows[3], "stepProgressRing4"));

                    var activeRing = FindControl<Control>(stepRows[4], "stepProgressRing5");
                    Assert.True(activeRing.Width >= 42, $"{activeRing.Name} should reserve enough width for the right-side progress ring.");
                    Assert.True(activeRing.Left > stepRows[4].Width - activeRing.Width - 24, $"{activeRing.Name} should be positioned on the right edge.");
                    Assert.True(IsProgressRingAnimated(activeRing), $"{activeRing.Name} should be animated while the step is active.");
                }
            });
        }

        [Fact]
        public void ProgressRingAdvancesAnimationFrame()
        {
            var dialogType = LoadAddInAssembly().GetType(
                "OfficeAgent.ExcelAddIn.Dialogs.BatchUploadProgressDialog",
                throwOnError: true);
            var ringType = dialogType.GetNestedType("StepProgressRing", BindingFlags.NonPublic)
                ?? throw new InvalidOperationException("StepProgressRing was not found.");
            var advanceFrame = ringType.GetMethod("AdvanceAnimationFrame", BindingFlags.NonPublic | BindingFlags.Instance)
                ?? throw new InvalidOperationException("StepProgressRing.AdvanceAnimationFrame was not found.");
            var startAngle = ringType.GetField("startAngle", BindingFlags.NonPublic | BindingFlags.Instance)
                ?? throw new InvalidOperationException("StepProgressRing.startAngle was not found.");
            var stateType = dialogType.GetNestedType("BatchUploadStepState", BindingFlags.NonPublic)
                ?? throw new InvalidOperationException("BatchUploadStepState was not found.");

            using (var ring = (Control)Activator.CreateInstance(ringType, Enum.Parse(stateType, "Active")))
            {
                var before = (int)startAngle.GetValue(ring);
                advanceFrame.Invoke(ring, Array.Empty<object>());
                var after = (int)startAngle.GetValue(ring);

                Assert.NotEqual(before, after);
            }
        }

        [Fact]
        public void DialogCanUpdateStepsAfterItIsCreated()
        {
            RunInSta(() =>
            {
                using (var dialog = CreateDialog())
                {
                    dialog.CreateControl();
                    dialog.PerformLayout();

                    InvokeDialogMethod(
                        dialog,
                        "SetStepActive",
                        2,
                        "字段验证中",
                        "正在重新验证字段",
                        "已开始字段验证");
                    dialog.PerformLayout();

                    var secondRow = FindControl<Control>(dialog, "stepRow2");
                    var fifthRow = FindControl<Control>(dialog, "stepRow5");
                    Assert.Equal("字段验证中", FindControl<Label>(secondRow, "stepTitleLabel2").Text);
                    Assert.Equal("正在重新验证字段", FindControl<Label>(secondRow, "stepDescriptionLabel2").Text);
                    Assert.Contains("已开始字段验证", FindControl<TextBox>(secondRow, "stepDetailsTextBox2").Text);
                    Assert.NotNull(TryFindControl<Control>(secondRow, "stepProgressRing2"));
                    Assert.Null(TryFindControl<Control>(fifthRow, "stepProgressRing5"));

                    InvokeDialogMethod(dialog, "AppendStepDetails", 2, "字段验证完成");
                    Assert.Contains("字段验证完成", FindControl<TextBox>(secondRow, "stepDetailsTextBox2").Text);

                    InvokeDialogMethod(
                        dialog,
                        "SetStepCompleted",
                        2,
                        "字段验证",
                        "验证通过",
                        "字段验证完成");
                    dialog.PerformLayout();

                    Assert.Equal("字段验证", FindControl<Label>(secondRow, "stepTitleLabel2").Text);
                    Assert.Equal("验证通过", FindControl<Label>(secondRow, "stepDescriptionLabel2").Text);
                    Assert.Null(TryFindControl<Control>(secondRow, "stepProgressRing2"));
                    Assert.Null(TryFindControl<TextBox>(secondRow, "stepDetailsTextBox2"));
                }
            });
        }

        private static Form CreateDialog()
        {
            return CreateDialog(null);
        }

        private static Form CreateDialog(string locale)
        {
            var type = LoadAddInAssembly().GetType(
                "OfficeAgent.ExcelAddIn.Dialogs.BatchUploadProgressDialog",
                throwOnError: true);
            if (string.IsNullOrEmpty(locale))
            {
                var createSample = type.GetMethod("CreateSample", BindingFlags.Public | BindingFlags.Static, null, Type.EmptyTypes, null)
                    ?? throw new InvalidOperationException("BatchUploadProgressDialog.CreateSample was not found.");

                return (Form)createSample.Invoke(null, Array.Empty<object>());
            }

            var createSampleWithLocale = type.GetMethod("CreateSample", BindingFlags.Public | BindingFlags.Static, null, new[] { typeof(string) }, null)
                ?? throw new InvalidOperationException("BatchUploadProgressDialog.CreateSample(locale) was not found.");

            return (Form)createSampleWithLocale.Invoke(null, new object[] { locale });
        }

        private static void InvokeDialogMethod(Form dialog, string methodName, params object[] arguments)
        {
            var method = dialog.GetType().GetMethod(methodName, BindingFlags.Public | BindingFlags.Instance)
                ?? throw new InvalidOperationException($"{methodName} was not found.");
            method.Invoke(dialog, arguments);
        }

        private static void ShowOffscreen(Form dialog)
        {
            dialog.StartPosition = FormStartPosition.Manual;
            dialog.Location = new Point(-32000, -32000);
            dialog.Show();
            Application.DoEvents();
            dialog.PerformLayout();
            Application.DoEvents();
        }

        private static Button[] VisibleFooterButtons(Form dialog)
        {
            return FindControl<Panel>(dialog, "footerPanel")
                .Controls
                .Cast<Control>()
                .OfType<Button>()
                .Where(button => button.Visible)
                .OrderBy(button => button.Left)
                .ToArray();
        }

        private static string BuildPreviewDetails()
        {
            return string.Join(
                "\r\n",
                new[]
                {
                    "部分上传将上传48个单元格，跳过4个单元格。",
                    "变更内容：",
                    "0331test / taskFlowNode_13882098334 -> 测试",
                    "0331test / taskFlowNode_13892195334 -> 1111",
                    "0331test / taskFlowNode_13883074334 -> 东努沙登加拉",
                    "0331test1 / SITEOWNER -> 15012344321",
                    "0331test1 / taskFlowNode_13882098334 -> 1111",
                    "0331test1 / taskFlowNode_13892195334 -> 1111",
                });
        }

        private static int CenterY(Control control)
        {
            return control.Top + (control.Height / 2);
        }

        private static string ResolveMarkerText(MethodInfo resolveText, Type stateType, string stateName, int stepNumber)
        {
            var state = Enum.Parse(stateType, stateName);
            return (string)resolveText.Invoke(null, new[] { state, stepNumber });
        }

        private static bool IsProgressRingAnimated(Control ring)
        {
            var animationTimer = ring.GetType().GetField("animationTimer", BindingFlags.NonPublic | BindingFlags.Instance)
                ?? throw new InvalidOperationException("StepProgressRing.animationTimer was not found.");
            return animationTimer.GetValue(ring) is System.Windows.Forms.Timer;
        }

        private static T FindControl<T>(Control root, string name)
            where T : Control
        {
            foreach (Control child in root.Controls)
            {
                if (child is T matched && string.Equals(child.Name, name, StringComparison.Ordinal))
                {
                    return matched;
                }

                var descendant = TryFindControl<T>(child, name);
                if (descendant != null)
                {
                    return descendant;
                }
            }

            throw new InvalidOperationException($"{typeof(T).Name} named '{name}' was not found.");
        }

        private static T TryFindControl<T>(Control root, string name)
            where T : Control
        {
            if (root is T matched && string.Equals(root.Name, name, StringComparison.Ordinal))
            {
                return matched;
            }

            foreach (Control child in root.Controls)
            {
                var descendant = TryFindControl<T>(child, name);
                if (descendant != null)
                {
                    return descendant;
                }
            }

            return null;
        }

        private static void AssertNoVisibleTextControlClips(Control root)
        {
            foreach (Control control in EnumerateControls(root))
            {
                if (!control.Visible || string.IsNullOrWhiteSpace(control.Text))
                {
                    continue;
                }

                if (control is Label label)
                {
                    var proposed = new Size(Math.Max(1, label.Width), int.MaxValue);
                    var measured = TextRenderer.MeasureText(
                        label.Text,
                        label.Font,
                        proposed,
                        TextFormatFlags.WordBreak | TextFormatFlags.TextBoxControl);
                    Assert.True(label.Height >= measured.Height || label.AutoSize, $"{label.Name} clips text '{label.Text}'.");
                }

                if (control is Button button)
                {
                    var measured = TextRenderer.MeasureText(button.Text, button.Font);
                    Assert.True(button.Width >= measured.Width + 12, $"{button.Name} clips button text '{button.Text}'.");
                    Assert.True(button.Height >= measured.Height + 6, $"{button.Name} clips button text '{button.Text}'.");
                }
            }
        }

        private static IEnumerable<Control> EnumerateControls(Control root)
        {
            foreach (Control child in root.Controls)
            {
                yield return child;
                foreach (var descendant in EnumerateControls(child))
                {
                    yield return descendant;
                }
            }
        }

        private static void ApplyFont(Control root, Font font)
        {
            root.Font = font;
            foreach (Control child in root.Controls)
            {
                ApplyFont(child, font);
            }
        }

        private static Assembly LoadAddInAssembly()
        {
            return Assembly.LoadFrom(ResolveAddInAssemblyPath());
        }

        private static string ResolveAddInAssemblyPath()
        {
            return System.IO.Path.GetFullPath(System.IO.Path.Combine(
                AppContext.BaseDirectory,
                "..",
                "..",
                "..",
                "..",
                "..",
                "src",
                "OfficeAgent.ExcelAddIn",
                "bin",
                "Debug",
                "OfficeAgent.ExcelAddIn.dll"));
        }

        private static void RunInSta(Action action)
        {
            Exception error = null;
            var thread = new Thread(() =>
            {
                try
                {
                    action();
                }
                catch (Exception ex)
                {
                    error = ex;
                }
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();

            if (error != null)
            {
                throw error;
            }
        }
    }
}
