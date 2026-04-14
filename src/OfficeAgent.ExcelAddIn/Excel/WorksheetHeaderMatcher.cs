using System;
using System.Collections.Generic;
using System.Linq;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Sync;

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal sealed class WorksheetHeaderMatcher
    {
        private readonly FieldMappingValueAccessor valueAccessor;

        public WorksheetHeaderMatcher(FieldMappingValueAccessor valueAccessor)
        {
            this.valueAccessor = valueAccessor ?? throw new ArgumentNullException(nameof(valueAccessor));
        }

        public WorksheetRuntimeColumn[] Match(
            string sheetName,
            SheetBinding binding,
            FieldMappingTableDefinition definition,
            IReadOnlyList<SheetFieldMappingRow> mappings,
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

            if (definition == null)
            {
                throw new ArgumentNullException(nameof(definition));
            }

            if (grid == null)
            {
                throw new ArgumentNullException(nameof(grid));
            }

            var rows = mappings ?? Array.Empty<SheetFieldMappingRow>();
            var result = new List<WorksheetRuntimeColumn>();
            var lastUsedColumn = grid.GetLastUsedColumn(sheetName);
            var headerRow = binding.HeaderStartRow <= 0 ? 1 : binding.HeaderStartRow;
            var currentParent = string.Empty;

            for (var column = 1; column <= lastUsedColumn; column++)
            {
                var topText = grid.GetCellText(sheetName, headerRow, column) ?? string.Empty;
                var bottomText = binding.HeaderRowCount > 1
                    ? grid.GetCellText(sheetName, headerRow + 1, column) ?? string.Empty
                    : string.Empty;

                if (!string.IsNullOrWhiteSpace(topText))
                {
                    currentParent = topText;
                }

                var match = FindMatch(definition, rows, topText, bottomText, currentParent, binding.HeaderRowCount);
                if (match == null)
                {
                    continue;
                }

                match.ColumnIndex = column;
                result.Add(match);
            }

            return result.ToArray();
        }

        private WorksheetRuntimeColumn FindMatch(
            FieldMappingTableDefinition definition,
            IReadOnlyList<SheetFieldMappingRow> mappings,
            string topText,
            string bottomText,
            string currentParent,
            int headerRowCount)
        {
            foreach (var mapping in mappings)
            {
                if (mapping == null)
                {
                    continue;
                }

                var headerType = valueAccessor.GetValue(definition, mapping, FieldMappingSemanticRole.HeaderType);
                var apiFieldKey = valueAccessor.GetValue(definition, mapping, FieldMappingSemanticRole.ApiFieldKey);
                var currentSingle = valueAccessor.GetValue(definition, mapping, FieldMappingSemanticRole.CurrentSingleHeaderText);
                var currentParentText = valueAccessor.GetValue(definition, mapping, FieldMappingSemanticRole.CurrentParentHeaderText);
                var currentChildText = valueAccessor.GetValue(definition, mapping, FieldMappingSemanticRole.CurrentChildHeaderText);

                if (headerRowCount <= 1)
                {
                    if (IsSingleHeader(headerType) &&
                        string.Equals(topText, currentSingle, StringComparison.Ordinal))
                    {
                        return new WorksheetRuntimeColumn
                        {
                            ApiFieldKey = apiFieldKey,
                            HeaderType = headerType,
                            DisplayText = currentSingle,
                            ParentDisplayText = string.Empty,
                            ChildDisplayText = string.Empty,
                            IsIdColumn = valueAccessor.GetBoolean(definition, mapping, FieldMappingSemanticRole.IsIdColumn),
                        };
                    }

                    continue;
                }

                if (IsSingleHeader(headerType) &&
                    string.Equals(topText, currentSingle, StringComparison.Ordinal) &&
                    (string.IsNullOrWhiteSpace(bottomText) || string.Equals(bottomText, topText, StringComparison.Ordinal)))
                {
                    return new WorksheetRuntimeColumn
                    {
                        ApiFieldKey = apiFieldKey,
                        HeaderType = headerType,
                        DisplayText = currentSingle,
                        ParentDisplayText = string.Empty,
                        ChildDisplayText = string.Empty,
                        IsIdColumn = valueAccessor.GetBoolean(definition, mapping, FieldMappingSemanticRole.IsIdColumn),
                    };
                }

                if (string.Equals(headerType, "activityProperty", StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(currentParent, currentParentText, StringComparison.Ordinal) &&
                    string.Equals(bottomText, currentChildText, StringComparison.Ordinal))
                {
                    return new WorksheetRuntimeColumn
                    {
                        ApiFieldKey = apiFieldKey,
                        HeaderType = headerType,
                        DisplayText = currentChildText,
                        ParentDisplayText = currentParentText,
                        ChildDisplayText = currentChildText,
                        IsIdColumn = valueAccessor.GetBoolean(definition, mapping, FieldMappingSemanticRole.IsIdColumn),
                    };
                }
            }

            return null;
        }

        private static bool IsSingleHeader(string headerType)
        {
            return string.IsNullOrWhiteSpace(headerType) ||
                   string.Equals(headerType, "single", StringComparison.OrdinalIgnoreCase);
        }
    }
}
