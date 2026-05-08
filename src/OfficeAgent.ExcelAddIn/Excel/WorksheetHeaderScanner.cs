using System;
using System.Collections.Generic;
using OfficeAgent.Core.Models;

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal sealed class WorksheetHeaderScanner
    {
        public AiColumnMappingActualHeader[] Scan(
            string sheetName,
            SheetBinding binding,
            IWorksheetGridAdapter grid)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                throw new ArgumentException("Sheet name is required.", nameof(sheetName));
            }

            if (binding == null)
            {
                throw new ArgumentNullException(nameof(binding));
            }

            if (grid == null)
            {
                throw new ArgumentNullException(nameof(grid));
            }

            var lastUsedColumn = grid.GetLastUsedColumn(sheetName);
            if (lastUsedColumn <= 0)
            {
                return Array.Empty<AiColumnMappingActualHeader>();
            }

            var headerRow = binding.HeaderStartRow <= 0 ? 1 : binding.HeaderStartRow;
            var headerRowCount = Math.Max(1, binding.HeaderRowCount);
            var result = new List<AiColumnMappingActualHeader>();
            var currentParent = string.Empty;

            for (var column = 1; column <= lastUsedColumn; column++)
            {
                var topText = grid.GetCellText(sheetName, headerRow, column) ?? string.Empty;
                var bottomText = headerRowCount > 1
                    ? grid.GetCellText(sheetName, headerRow + 1, column) ?? string.Empty
                    : string.Empty;

                if (!string.IsNullOrWhiteSpace(topText))
                {
                    currentParent = topText;
                }

                var actualL1 = headerRowCount > 1 && string.IsNullOrWhiteSpace(topText) && !string.IsNullOrWhiteSpace(bottomText)
                    ? currentParent
                    : topText;
                var actualL2 = headerRowCount > 1 && !IsMergedSingleHeader(topText, bottomText)
                    ? bottomText
                    : string.Empty;

                if (string.IsNullOrWhiteSpace(actualL1) && string.IsNullOrWhiteSpace(actualL2))
                {
                    continue;
                }

                result.Add(new AiColumnMappingActualHeader
                {
                    ExcelColumn = column,
                    ActualL1 = actualL1 ?? string.Empty,
                    ActualL2 = actualL2 ?? string.Empty,
                    DisplayText = FormatDisplayText(actualL1, actualL2),
                });
            }

            return result.ToArray();
        }

        private static bool IsMergedSingleHeader(string topText, string bottomText)
        {
            return string.IsNullOrWhiteSpace(bottomText) ||
                   string.Equals(topText, bottomText, StringComparison.Ordinal);
        }

        private static string FormatDisplayText(string actualL1, string actualL2)
        {
            if (string.IsNullOrWhiteSpace(actualL2))
            {
                return actualL1 ?? string.Empty;
            }

            return (actualL1 ?? string.Empty) + "/" + actualL2;
        }
    }
}
