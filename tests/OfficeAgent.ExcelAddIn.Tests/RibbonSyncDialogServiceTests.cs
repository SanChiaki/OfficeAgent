using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using OfficeAgent.Core.Models;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class RibbonSyncDialogServiceTests
    {
        [Fact]
        public void ShowProjectLayoutDialogReturnsNullWhenUserCancels()
        {
            var seed = CreateSeedBinding();
            var result = InvokeShowProjectLayoutDialog(
                seed,
                dialog =>
                {
                    dialog.DialogResult = DialogResult.Cancel;
                    dialog.Close();
                });

            Assert.Null(result);
        }

        [Fact]
        public void ShowProjectLayoutDialogReturnsEditedBindingWhenUserConfirms()
        {
            var seed = CreateSeedBinding();
            var result = InvokeShowProjectLayoutDialog(
                seed,
                dialog =>
                {
                    GetRequiredTextBox(dialog, "HeaderStartRowTextBox").Text = "4";
                    GetRequiredTextBox(dialog, "HeaderRowCountTextBox").Text = "1";
                    GetRequiredTextBox(dialog, "DataStartRowTextBox").Text = "5";

                    var okButton = dialog.Controls
                        .OfType<Button>()
                        .Single(button => string.Equals(button.Text, "确定", StringComparison.Ordinal));
                    okButton.PerformClick();
                });

            var binding = Assert.IsType<SheetBinding>(result);
            Assert.Equal("Sheet1", binding.SheetName);
            Assert.Equal("current-business-system", binding.SystemKey);
            Assert.Equal("performance", binding.ProjectId);
            Assert.Equal("绩效项目", binding.ProjectName);
            Assert.Equal(4, binding.HeaderStartRow);
            Assert.Equal(1, binding.HeaderRowCount);
            Assert.Equal(5, binding.DataStartRow);
        }

        private static object InvokeShowProjectLayoutDialog(SheetBinding suggestedBinding, Action<Form> automateDialog)
        {
            object returnValue = null;
            Exception capturedException = null;
            using (var completed = new ManualResetEventSlim(false))
            {
                var thread = new Thread(
                    () =>
                    {
                        try
                        {
                            var addInAssembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
                            var serviceType = addInAssembly.GetType("OfficeAgent.ExcelAddIn.Dialogs.RibbonSyncDialogService", throwOnError: true);
                            var dialogType = addInAssembly.GetType("OfficeAgent.ExcelAddIn.Dialogs.ProjectLayoutDialog", throwOnError: true);
                            var service = Activator.CreateInstance(serviceType, nonPublic: true);

                            var timer = new System.Windows.Forms.Timer { Interval = 25 };
                            timer.Tick += (sender, args) =>
                            {
                                var dialog = Application.OpenForms
                                    .Cast<Form>()
                                    .FirstOrDefault(form => dialogType.IsInstanceOfType(form));
                                if (dialog == null)
                                {
                                    return;
                                }

                                timer.Stop();
                                automateDialog(dialog);
                            };
                            timer.Start();

                            returnValue = serviceType
                                .GetMethod("ShowProjectLayoutDialog", BindingFlags.Instance | BindingFlags.Public)
                                .Invoke(service, new object[] { suggestedBinding });
                        }
                        catch (Exception ex)
                        {
                            capturedException = ex;
                        }
                        finally
                        {
                            completed.Set();
                        }
                    });

                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();

                if (!completed.Wait(TimeSpan.FromSeconds(5)))
                {
                    throw new TimeoutException("Timed out waiting for dialog interaction test to complete.");
                }
            }

            if (capturedException != null)
            {
                throw capturedException;
            }

            return returnValue;
        }

        private static TextBox GetRequiredTextBox(Form dialog, string controlName)
        {
            var matched = dialog.Controls.Find(controlName, searchAllChildren: true);
            var textBox = Assert.Single(matched);
            return Assert.IsType<TextBox>(textBox);
        }

        private static SheetBinding CreateSeedBinding()
        {
            return new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 1,
                HeaderRowCount = 2,
                DataStartRow = 3,
            };
        }

        private static string ResolveAddInAssemblyPath()
        {
            return Path.GetFullPath(
                Path.Combine(
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
    }
}
