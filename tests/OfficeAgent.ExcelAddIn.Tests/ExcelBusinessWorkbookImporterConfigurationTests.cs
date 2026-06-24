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

            Assert.Contains("private const string BusinessDataSheetName = \"Business Data\";", text, StringComparison.Ordinal);
            Assert.Contains(
                "FindWorksheet(sourceWorkbook, BusinessDataSheetName, StringComparison.Ordinal)",
                text,
                StringComparison.Ordinal);
            Assert.DoesNotContain("FindWorksheet(sourceWorkbook, \"Business Data\"", text, StringComparison.Ordinal);
            Assert.Contains("var originalTargetSheetName = targetWorksheet.Name;", text, StringComparison.Ordinal);
            Assert.Contains("targetWorksheet.Name = originalTargetSheetName;", text, StringComparison.Ordinal);
        }

        [Fact]
        public void ImporterDeletesTemporaryWorkbookInFinallyBlock()
        {
            var text = ReadSource("src", "OfficeAgent.ExcelAddIn", "Excel", "ExcelBusinessWorkbookImporter.cs");
            var importMethodBody = ExtractMethodBody(text, "ImportBusinessDataSheet");
            var finallyBlock = ExtractBlockAfterKeyword(importMethodBody, "finally");

            Assert.Contains("File.Delete(tempPath);", finallyBlock, StringComparison.Ordinal);
        }

        [Fact]
        public void ImporterDetectsContentWithConstantsAndFormulas()
        {
            var text = ReadSource("src", "OfficeAgent.ExcelAddIn", "Excel", "ExcelBusinessWorkbookImporter.cs");
            var methodBody = ExtractMethodBody(text, "IsWorkSheetContentBlank");

            Assert.Contains(
                "HasSpecialCells(worksheet, ExcelInterop.XlCellType.xlCellTypeConstants)",
                methodBody,
                StringComparison.Ordinal);
            Assert.Contains(
                "HasSpecialCells(worksheet, ExcelInterop.XlCellType.xlCellTypeFormulas)",
                methodBody,
                StringComparison.Ordinal);
            Assert.DoesNotContain("UsedRange", methodBody, StringComparison.Ordinal);
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

        private static string ExtractMethodBody(string text, string methodName)
        {
            var methodIndex = text.IndexOf(methodName, StringComparison.Ordinal);
            Assert.True(methodIndex >= 0, $"Method '{methodName}' was not found.");

            var openBraceIndex = text.IndexOf('{', methodIndex);
            Assert.True(openBraceIndex >= 0, $"Method '{methodName}' did not have a body.");

            return ExtractBraceBlock(text, openBraceIndex);
        }

        private static string ExtractBlockAfterKeyword(string text, string keyword)
        {
            var keywordIndex = text.IndexOf(keyword, StringComparison.Ordinal);
            Assert.True(keywordIndex >= 0, $"Keyword '{keyword}' was not found.");

            var openBraceIndex = text.IndexOf('{', keywordIndex);
            Assert.True(openBraceIndex >= 0, $"Keyword '{keyword}' did not have a block.");

            return ExtractBraceBlock(text, openBraceIndex);
        }

        private static string ExtractBraceBlock(string text, int openBraceIndex)
        {
            var depth = 0;
            for (var index = openBraceIndex; index < text.Length; index++)
            {
                if (text[index] == '{')
                {
                    depth++;
                    continue;
                }

                if (text[index] != '}')
                {
                    continue;
                }

                depth--;
                if (depth == 0)
                {
                    return text.Substring(openBraceIndex, index - openBraceIndex + 1);
                }
            }

            throw new InvalidOperationException("Brace block was not closed.");
        }
    }
}
