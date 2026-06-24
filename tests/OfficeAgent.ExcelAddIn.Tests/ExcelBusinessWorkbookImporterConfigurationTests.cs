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
            var deleteIndex = finallyBlock.IndexOf("File.Delete(tempPath);", StringComparison.Ordinal);
            Assert.True(deleteIndex >= 0, "Temporary workbook delete was not found in the import finally block.");
            var deleteGuard = ExtractEnclosingBlock(finallyBlock, deleteIndex, "try");

            Assert.Contains("File.Delete(tempPath);", deleteGuard, StringComparison.Ordinal);
            Assert.Contains("catch", finallyBlock.Substring(deleteIndex), StringComparison.Ordinal);
        }

        [Fact]
        public void ImporterRestoresTargetWorksheetAfterBestEffortFreezePaneCopy()
        {
            var text = ReadSource("src", "OfficeAgent.ExcelAddIn", "Excel", "ExcelBusinessWorkbookImporter.cs");
            var methodBody = ExtractMethodBody(text, "TryCopyFreezePaneState");
            var sourceActivateIndex = methodBody.IndexOf("sourceWorksheet.Activate();", StringComparison.Ordinal);
            Assert.True(sourceActivateIndex >= 0, "Freeze pane copy should activate the source worksheet before reading pane state.");
            var finallyBlock = ExtractBlockAfterKeyword(methodBody, "finally");
            var targetActivateIndex = finallyBlock.IndexOf("targetWorksheet.Activate();", StringComparison.Ordinal);

            Assert.True(targetActivateIndex >= 0, "Freeze pane copy should restore the target worksheet in finally.");
            Assert.True(
                methodBody.IndexOf("finally", StringComparison.Ordinal) > sourceActivateIndex,
                "Target worksheet restoration should occur after source worksheet activation.");
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
            var methodIndex = FindMethodDeclarationIndex(text, methodName);
            Assert.True(methodIndex >= 0, $"Method '{methodName}' was not found.");

            var openBraceIndex = text.IndexOf('{', methodIndex);
            Assert.True(openBraceIndex >= 0, $"Method '{methodName}' did not have a body.");

            return ExtractBraceBlock(text, openBraceIndex);
        }

        private static int FindMethodDeclarationIndex(string text, string methodName)
        {
            var searchIndex = 0;
            while (searchIndex < text.Length)
            {
                var methodIndex = text.IndexOf(methodName, searchIndex, StringComparison.Ordinal);
                if (methodIndex < 0)
                {
                    return -1;
                }

                var lineStart = text.LastIndexOf('\n', methodIndex);
                lineStart = lineStart < 0 ? 0 : lineStart + 1;
                var line = text.Substring(lineStart, methodIndex - lineStart);
                if (line.Contains("public ") ||
                    line.Contains("private ") ||
                    line.Contains("internal "))
                {
                    return methodIndex;
                }

                searchIndex = methodIndex + methodName.Length;
            }

            return -1;
        }

        private static string ExtractBlockAfterKeyword(string text, string keyword)
        {
            var keywordIndex = text.IndexOf(keyword, StringComparison.Ordinal);
            Assert.True(keywordIndex >= 0, $"Keyword '{keyword}' was not found.");

            var openBraceIndex = text.IndexOf('{', keywordIndex);
            Assert.True(openBraceIndex >= 0, $"Keyword '{keyword}' did not have a block.");

            return ExtractBraceBlock(text, openBraceIndex);
        }

        private static string ExtractEnclosingBlock(string text, int innerIndex, string keyword)
        {
            var searchIndex = innerIndex;
            while (searchIndex >= 0)
            {
                var keywordIndex = text.LastIndexOf(keyword, searchIndex, StringComparison.Ordinal);
                Assert.True(keywordIndex >= 0, $"No enclosing '{keyword}' block was found.");

                var openBraceIndex = text.IndexOf('{', keywordIndex);
                Assert.True(openBraceIndex >= 0, $"Keyword '{keyword}' did not have a block.");

                var block = ExtractBraceBlock(text, openBraceIndex);
                if (openBraceIndex <= innerIndex && innerIndex < openBraceIndex + block.Length)
                {
                    return block;
                }

                searchIndex = keywordIndex - 1;
            }

            throw new InvalidOperationException($"No enclosing '{keyword}' block was found.");
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
