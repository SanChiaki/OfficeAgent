using System;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class TemplateDialogConfigurationTests
    {
        [Fact]
        public void TemplatePickerDialogUsesWrappingInstructionLayout()
        {
            var dialogText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "Dialogs",
                "TemplatePickerDialog.cs"));

            Assert.Contains("TableLayoutPanel", dialogText, StringComparison.Ordinal);
            Assert.Contains("AutoSize = true", dialogText, StringComparison.Ordinal);
            Assert.Contains("MaximumSize = new Size(", dialogText, StringComparison.Ordinal);
            Assert.DoesNotContain("Height = 48", dialogText, StringComparison.Ordinal);
        }

        [Fact]
        public void TemplateNameDialogUsesWrappingPromptAndDedicatedContentPanel()
        {
            var dialogText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "Dialogs",
                "TemplateNameDialog.cs"));

            Assert.Contains("TableLayoutPanel", dialogText, StringComparison.Ordinal);
            Assert.Contains("AutoSize = true", dialogText, StringComparison.Ordinal);
            Assert.Contains("MaximumSize = new Size(", dialogText, StringComparison.Ordinal);
            Assert.Contains("Padding = new Padding(16", dialogText, StringComparison.Ordinal);
        }

        [Fact]
        public void TemplateDialogServiceUsesDedicatedTemplateConfirmationDialogs()
        {
            var dialogServiceText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "Dialogs",
                "TemplateDialogService.cs"));

            Assert.Contains("TemplateOverwriteConfirmDialog.Confirm(", dialogServiceText, StringComparison.Ordinal);
            Assert.Contains("TemplateRevisionConflictDialog.ShowDecision(", dialogServiceText, StringComparison.Ordinal);
            Assert.DoesNotContain("MessageBox.Show(", dialogServiceText, StringComparison.Ordinal);
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
