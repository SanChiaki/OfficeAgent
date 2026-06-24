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
