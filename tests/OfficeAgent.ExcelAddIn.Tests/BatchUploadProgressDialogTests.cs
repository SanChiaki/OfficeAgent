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
                    var previewTextBox = FindControl<TextBox>(dialog, "stepDetailsTextBox3");

                    Assert.DoesNotContain(EnumerateControls(dialog), control => string.Equals(control.Name, "closeGlyphButton", StringComparison.Ordinal));
                    Assert.True(contentPanel.AutoScroll);
                    Assert.Equal(DockStyle.Fill, contentPanel.Dock);
                    Assert.Equal(DockStyle.Bottom, footerPanel.Dock);
                    Assert.Equal(FlowDirection.TopDown, stepsPanel.FlowDirection);
                    Assert.True(stepsPanel.AutoSize);
                    Assert.True(previewTextBox.Multiline);
                    Assert.Equal(ScrollBars.Vertical, previewTextBox.ScrollBars);
                    Assert.True(previewTextBox.Width >= contentPanel.ClientSize.Width - 160);

                    AssertNoVisibleTextControlClips(dialog);
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

        private static Form CreateDialog()
        {
            var type = LoadAddInAssembly().GetType(
                "OfficeAgent.ExcelAddIn.Dialogs.BatchUploadProgressDialog",
                throwOnError: true);
            var createSample = type.GetMethod("CreateSample", BindingFlags.Public | BindingFlags.Static)
                ?? throw new InvalidOperationException("BatchUploadProgressDialog.CreateSample was not found.");

            return (Form)createSample.Invoke(null, Array.Empty<object>());
        }

        private static string ResolveMarkerText(MethodInfo resolveText, Type stateType, string stateName, int stepNumber)
        {
            var state = Enum.Parse(stateType, stateName);
            return (string)resolveText.Invoke(null, new[] { state, stepNumber });
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
