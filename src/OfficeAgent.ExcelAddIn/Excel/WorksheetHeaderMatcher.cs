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

            var rows = (mappings ?? Array.Empty<SheetFieldMappingRow>())
                .Where(mapping => mapping != null && string.Equals(mapping.SheetName, sheetName, StringComparison.OrdinalIgnoreCase))
                .ToArray();
            var lookup = BuildLookup(definition, rows);
            if (binding.HeaderRowCount == 1 && lookup.HasGroupedSingleMetadata)
            {
                throw new InvalidOperationException($"当前 HeaderRowCount=1，无法识别带 Excel L2 的 single 表头，请先修正 {MetadataWorksheetNames.Current}。");
            }

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

                var match = FindMatch(lookup, topText, bottomText, currentParent, binding.HeaderRowCount);
                if (match == null)
                {
                    continue;
                }

                match.ColumnIndex = column;
                result.Add(match);
            }

            return result.ToArray();
        }

        private HeaderLookup BuildLookup(
            FieldMappingTableDefinition definition,
            IReadOnlyList<SheetFieldMappingRow> mappings)
        {
            var singleHeaders = new Dictionary<string, WorksheetRuntimeColumn>(StringComparer.Ordinal);
            var groupedSingleHeaders = new Dictionary<string, WorksheetRuntimeColumn>(StringComparer.Ordinal);
            var singleRowActivityHeaders = new Dictionary<string, WorksheetRuntimeColumn>(StringComparer.Ordinal);
            var activityHeaders = new Dictionary<string, WorksheetRuntimeColumn>(StringComparer.Ordinal);
            var hasGroupedSingleMetadata = false;

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
                var template = new WorksheetRuntimeColumn
                {
                    ApiFieldKey = apiFieldKey,
                    HeaderType = headerType,
                    DisplayText = currentSingle,
                    ParentDisplayText = currentParentText,
                    ChildDisplayText = currentChildText,
                    IsIdColumn = valueAccessor.GetBoolean(definition, mapping, FieldMappingSemanticRole.IsIdColumn),
                };

                if (IsSingleHeader(headerType))
                {
                    if (!string.IsNullOrWhiteSpace(currentChildText))
                    {
                        hasGroupedSingleMetadata = true;
                        var groupedKey = BuildTwoLevelKey(currentParentText, currentChildText);
                        if (!string.IsNullOrWhiteSpace(groupedKey))
                        {
                            if (groupedSingleHeaders.ContainsKey(groupedKey) || activityHeaders.ContainsKey(groupedKey))
                            {
                                throw new InvalidOperationException($"SheetFieldMappings 中存在重复的双层表头键，请先修正 {MetadataWorksheetNames.Current}。");
                            }

                            groupedSingleHeaders[groupedKey] = template;
                        }
                    }
                    else if (!string.IsNullOrWhiteSpace(currentSingle) && !singleHeaders.ContainsKey(currentSingle))
                    {
                        singleHeaders[currentSingle] = template;
                    }

                    continue;
                }

                if (string.Equals(headerType, "activityProperty", StringComparison.OrdinalIgnoreCase))
                {
                    AddHeader(
                        singleRowActivityHeaders,
                        ResolveSingleRowActivityDisplayText(template),
                        template);

                    var activityKey = BuildTwoLevelKey(currentParentText, currentChildText);
                    if (!string.IsNullOrWhiteSpace(activityKey))
                    {
                        if (groupedSingleHeaders.ContainsKey(activityKey))
                        {
                            throw new InvalidOperationException($"SheetFieldMappings 中存在重复的双层表头键，请先修正 {MetadataWorksheetNames.Current}。");
                        }

                        if (!activityHeaders.ContainsKey(activityKey))
                        {
                            activityHeaders[activityKey] = template;
                        }
                    }
                }
            }

            return new HeaderLookup(singleHeaders, groupedSingleHeaders, singleRowActivityHeaders, activityHeaders, hasGroupedSingleMetadata);
        }

        private WorksheetRuntimeColumn FindMatch(
            HeaderLookup lookup,
            string topText,
            string bottomText,
            string currentParent,
            int headerRowCount)
        {
            if (headerRowCount <= 1)
            {
                if (lookup.SingleHeaders.TryGetValue(topText, out var singleHeader))
                {
                    return CloneSingleHeader(singleHeader);
                }

                if (lookup.SingleRowActivityHeaders.TryGetValue(topText, out var singleRowActivityHeader))
                {
                    return CloneSingleRowActivityHeader(singleRowActivityHeader, topText);
                }

                return null;
            }

            if (lookup.SingleHeaders.TryGetValue(topText, out var mergedSingleHeader) &&
                (string.IsNullOrWhiteSpace(bottomText) || string.Equals(bottomText, topText, StringComparison.Ordinal)))
            {
                return CloneSingleHeader(mergedSingleHeader);
            }

            var activityLookupKey = BuildTwoLevelKey(currentParent, bottomText);
            if (lookup.GroupedSingleHeaders.TryGetValue(activityLookupKey, out var groupedSingleHeader))
            {
                return CloneGroupedSingleHeader(groupedSingleHeader);
            }

            if (lookup.ActivityHeaders.TryGetValue(activityLookupKey, out var activityHeader))
            {
                return CloneActivityHeader(activityHeader);
            }

            return null;
        }

        private static WorksheetRuntimeColumn CloneSingleHeader(WorksheetRuntimeColumn template)
        {
            return new WorksheetRuntimeColumn
            {
                ApiFieldKey = template.ApiFieldKey,
                HeaderType = template.HeaderType,
                DisplayText = template.DisplayText,
                ParentDisplayText = string.Empty,
                ChildDisplayText = string.Empty,
                IsIdColumn = template.IsIdColumn,
            };
        }

        private static WorksheetRuntimeColumn CloneActivityHeader(WorksheetRuntimeColumn template)
        {
            return new WorksheetRuntimeColumn
            {
                ApiFieldKey = template.ApiFieldKey,
                HeaderType = template.HeaderType,
                DisplayText = template.ChildDisplayText,
                ParentDisplayText = template.ParentDisplayText,
                ChildDisplayText = template.ChildDisplayText,
                IsIdColumn = template.IsIdColumn,
            };
        }

        private static WorksheetRuntimeColumn CloneSingleRowActivityHeader(WorksheetRuntimeColumn template, string displayText)
        {
            var resolvedDisplayText = string.IsNullOrWhiteSpace(displayText)
                ? ResolveSingleRowActivityDisplayText(template)
                : displayText;

            return new WorksheetRuntimeColumn
            {
                ApiFieldKey = template.ApiFieldKey,
                HeaderType = template.HeaderType,
                DisplayText = resolvedDisplayText,
                ParentDisplayText = string.Empty,
                ChildDisplayText = resolvedDisplayText,
                IsIdColumn = template.IsIdColumn,
            };
        }

        private static WorksheetRuntimeColumn CloneGroupedSingleHeader(WorksheetRuntimeColumn template)
        {
            return new WorksheetRuntimeColumn
            {
                ApiFieldKey = template.ApiFieldKey,
                HeaderType = template.HeaderType,
                DisplayText = template.ChildDisplayText,
                ParentDisplayText = template.ParentDisplayText,
                ChildDisplayText = template.ChildDisplayText,
                IsIdColumn = template.IsIdColumn,
            };
        }

        private static void AddHeader(
            IDictionary<string, WorksheetRuntimeColumn> headers,
            string headerText,
            WorksheetRuntimeColumn template)
        {
            if (string.IsNullOrWhiteSpace(headerText) || headers.ContainsKey(headerText))
            {
                return;
            }

            headers[headerText] = template;
        }

        private static string ResolveSingleRowActivityDisplayText(WorksheetRuntimeColumn template)
        {
            if (template == null)
            {
                return string.Empty;
            }

            return !string.IsNullOrWhiteSpace(template.ParentDisplayText)
                ? template.ParentDisplayText
                : template.ChildDisplayText ?? string.Empty;
        }

        private static string BuildTwoLevelKey(string parentText, string childText)
        {
            if (string.IsNullOrWhiteSpace(parentText) || string.IsNullOrWhiteSpace(childText))
            {
                return string.Empty;
            }

            return parentText + "\u001f" + childText;
        }

        private sealed class HeaderLookup
        {
            public HeaderLookup(
                IReadOnlyDictionary<string, WorksheetRuntimeColumn> singleHeaders,
                IReadOnlyDictionary<string, WorksheetRuntimeColumn> groupedSingleHeaders,
                IReadOnlyDictionary<string, WorksheetRuntimeColumn> singleRowActivityHeaders,
                IReadOnlyDictionary<string, WorksheetRuntimeColumn> activityHeaders,
                bool hasGroupedSingleMetadata)
            {
                SingleHeaders = singleHeaders ?? throw new ArgumentNullException(nameof(singleHeaders));
                GroupedSingleHeaders = groupedSingleHeaders ?? throw new ArgumentNullException(nameof(groupedSingleHeaders));
                SingleRowActivityHeaders = singleRowActivityHeaders ?? throw new ArgumentNullException(nameof(singleRowActivityHeaders));
                ActivityHeaders = activityHeaders ?? throw new ArgumentNullException(nameof(activityHeaders));
                HasGroupedSingleMetadata = hasGroupedSingleMetadata;
            }

            public IReadOnlyDictionary<string, WorksheetRuntimeColumn> SingleHeaders { get; }

            public IReadOnlyDictionary<string, WorksheetRuntimeColumn> GroupedSingleHeaders { get; }

            public IReadOnlyDictionary<string, WorksheetRuntimeColumn> SingleRowActivityHeaders { get; }

            public IReadOnlyDictionary<string, WorksheetRuntimeColumn> ActivityHeaders { get; }

            public bool HasGroupedSingleMetadata { get; }
        }

        private static bool IsSingleHeader(string headerType)
        {
            return string.IsNullOrWhiteSpace(headerType) ||
                   string.Equals(headerType, "single", StringComparison.OrdinalIgnoreCase);
        }
    }
}
