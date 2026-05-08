using System;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class AiColumnMappingDialogOwnerTests
    {
        [Fact]
        public void RibbonSyncDialogServiceUsesExcelWindowOwnerForAiMappingDialogs()
        {
            var dialogServiceText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "Dialogs",
                "OperationResultDialog.cs"));

            Assert.Contains("ExcelDialogOwner.FromCurrentApplication()", dialogServiceText, StringComparison.Ordinal);
            Assert.Contains("AiColumnMappingPreviewDialog.Confirm(preview, owner)", dialogServiceText, StringComparison.Ordinal);
            Assert.Contains("AiColumnMappingProgressDialog.Run(owner", dialogServiceText, StringComparison.Ordinal);
        }

        [Fact]
        public void AiColumnMappingPreviewDialogShowsWithOwnerWindow()
        {
            var dialogText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "Dialogs",
                "AiColumnMappingPreviewDialog.cs"));

            Assert.Contains("Confirm(AiColumnMappingPreview preview, IWin32Window owner)", dialogText, StringComparison.Ordinal);
            Assert.Contains("dialog.ShowDialog(owner)", dialogText, StringComparison.Ordinal);
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
