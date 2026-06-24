using System;
using System.IO;
using System.Linq;
using System.Reflection;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class InitializeSheetDialogTests
    {
        [Theory]
        [InlineData(true, true, "TemplateImport")]
        [InlineData(true, false, "ConfigOnly")]
        [InlineData(false, true, "ConfigOnly")]
        [InlineData(false, false, "ConfigOnly")]
        public void ResolveDefaultModeFollowsBlankSheetPolicy(
            bool isBlankSheet,
            bool canImportTemplate,
            string expectedModeName)
        {
            var dialogType = LoadAddInAssembly().GetType(
                "OfficeAgent.ExcelAddIn.Dialogs.InitializeSheetDialog",
                throwOnError: true);
            var method = dialogType.GetMethod(
                "ResolveDefaultMode",
                BindingFlags.Static | BindingFlags.NonPublic);

            Assert.NotNull(method);

            var result = method.Invoke(null, new object[] { isBlankSheet, canImportTemplate });

            Assert.Equal(expectedModeName, result.ToString());
        }

        [Fact]
        public void OverwriteRiskTextIsResolvedFromHostLocalizedStrings()
        {
            var dialogSourcePath = ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "Dialogs",
                "InitializeSheetDialog.cs");

            Assert.True(File.Exists(dialogSourcePath), "InitializeSheetDialog.cs should exist.");

            var dialogSource = File.ReadAllText(dialogSourcePath);

            Assert.Contains("InitializeSheetOverwriteRiskMessage", dialogSource, StringComparison.Ordinal);
            Assert.DoesNotContain("覆盖当前表", dialogSource, StringComparison.Ordinal);
            Assert.DoesNotContain("模板创建作业表会覆盖", dialogSource, StringComparison.Ordinal);
        }

        [Fact]
        public void TemplateLoadingRunsFromShownOnBackgroundTaskWithConfirmDisabledWhileLoading()
        {
            var dialogSourcePath = ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "Dialogs",
                "InitializeSheetDialog.cs");

            Assert.True(File.Exists(dialogSourcePath), "InitializeSheetDialog.cs should exist.");

            var dialogSource = File.ReadAllText(dialogSourcePath);
            var disableConfirmIndex = dialogSource.IndexOf("confirmButton.Enabled = false;", StringComparison.Ordinal);
            var backgroundLoadIndex = dialogSource.IndexOf("Task.Run(loadTemplates)", StringComparison.Ordinal);

            Assert.Contains("protected override async void OnShown", dialogSource, StringComparison.Ordinal);
            Assert.True(disableConfirmIndex >= 0, "Confirm should be explicitly disabled when template loading starts.");
            Assert.True(backgroundLoadIndex >= 0, "Template loading should run on a background task from OnShown.");
            Assert.True(
                disableConfirmIndex < backgroundLoadIndex,
                "Confirm should be disabled before the background template load starts.");
        }

        [Fact]
        public void EmptyTemplateFallbackUsesInitializeSheetSpecificMessage()
        {
            var dialogSourcePath = ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "Dialogs",
                "InitializeSheetDialog.cs");

            Assert.True(File.Exists(dialogSourcePath), "InitializeSheetDialog.cs should exist.");

            var dialogSource = File.ReadAllText(dialogSourcePath);

            Assert.Contains("InitializeSheetTemplateEmptyMessage", dialogSource, StringComparison.Ordinal);
            Assert.DoesNotContain("strings.TemplateNoAvailableMessage", dialogSource, StringComparison.Ordinal);
        }

        [Fact]
        public void ImportProgressTreatsOperationCanceledExceptionAsUserCancelOnlyWhenTokenSourceWasCanceled()
        {
            var dialogSourcePath = ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "Dialogs",
                "InitializeSheetImportProgressDialog.cs");

            Assert.True(File.Exists(dialogSourcePath), "InitializeSheetImportProgressDialog.cs should exist.");

            var dialogSource = File.ReadAllText(dialogSourcePath);
            var cancellationGuardIndex = dialogSource.IndexOf(
                "if (cancellationTokenSource.IsCancellationRequested)",
                StringComparison.Ordinal);
            var userCanceledAssignmentIndex = dialogSource.IndexOf("canceled = true;", StringComparison.Ordinal);
            var errorAssignmentIndex = dialogSource.IndexOf("error = ex;", StringComparison.Ordinal);

            Assert.Contains("catch (OperationCanceledException ex)", dialogSource, StringComparison.Ordinal);
            Assert.True(cancellationGuardIndex >= 0, "OperationCanceledException should be classified by the dialog's token source.");
            Assert.True(userCanceledAssignmentIndex > cancellationGuardIndex, "User cancellation should only be set inside the token-source guard.");
            Assert.True(errorAssignmentIndex > userCanceledAssignmentIndex, "Non-user cancellation should be stored as an error.");
        }

        private static Assembly LoadAddInAssembly()
        {
            return Assembly.LoadFrom(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "bin",
                "Debug",
                "OfficeAgent.ExcelAddIn.dll"));
        }

        private static string ResolveRepositoryPath(params string[] segments)
        {
            return Path.GetFullPath(Path.Combine(new[]
            {
                AppContext.BaseDirectory,
                "..",
                "..",
                "..",
                "..",
                "..",
            }.Concat(segments).ToArray()));
        }
    }
}
