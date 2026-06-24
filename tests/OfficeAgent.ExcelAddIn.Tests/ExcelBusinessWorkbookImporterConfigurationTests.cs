using System;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class ExcelBusinessWorkbookImporterConfigurationTests
    {
        [Fact]
        public void ImporterUsesBusinessDataSheetNameAndPreservesTargetSheetName()
        {
            var text = ReadSource("src", "OfficeAgent.ExcelAddIn", "Excel", "ExcelBusinessWorkbookImporter.cs");

            Assert.Contains("Business Data", text, StringComparison.Ordinal);
            Assert.Contains("var originalTargetSheetName = targetWorksheet.Name;", text, StringComparison.Ordinal);
            Assert.Contains("targetWorksheet.Name = originalTargetSheetName;", text, StringComparison.Ordinal);
        }

        [Fact]
        public void ImporterDeletesTemporaryWorkbookInFinallyBlock()
        {
            var text = ReadSource("src", "OfficeAgent.ExcelAddIn", "Excel", "ExcelBusinessWorkbookImporter.cs");

            Assert.Contains("finally", text, StringComparison.Ordinal);
            Assert.Contains("File.Delete(tempPath);", text, StringComparison.Ordinal);
        }

        [Fact]
        public void ImporterDetectsContentWithConstantsAndFormulas()
        {
            var text = ReadSource("src", "OfficeAgent.ExcelAddIn", "Excel", "ExcelBusinessWorkbookImporter.cs");

            Assert.Contains("xlCellTypeConstants", text, StringComparison.Ordinal);
            Assert.Contains("xlCellTypeFormulas", text, StringComparison.Ordinal);
        }

        private static string ReadSource(params string[] segments)
        {
            return File.ReadAllText(ResolveRepositoryPath(segments));
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
