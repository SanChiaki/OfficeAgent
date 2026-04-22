using System;
using System.Collections.Generic;
using System.Linq;
using OfficeAgent.Core.Models;

namespace OfficeAgent.Infrastructure.Http
{
    public static class CurrentBusinessFieldMappingColumns
    {
        public const string HeaderType = "HeaderType";
        public const string DefaultLevel1 = "DefaultL1";
        public const string CurrentLevel1 = "CurrentL1";
        public const string DefaultLevel2 = "DefaultL2";
        public const string CurrentLevel2 = "CurrentL2";
        public const string HeaderId = "HeaderId";
        public const string ApiFieldKey = "ApiFieldKey";
        public const string IsIdColumn = "IsIdColumn";
        public const string ActivityId = "ActivityId";
        public const string PropertyId = "PropertyId";
        public const string SingleHeaderType = "single";
        public const string ActivityPropertyHeaderType = "activityProperty";
    }

    public static class CurrentBusinessFieldMappingHeaders
    {
        public const string HeaderType = "HeaderType";
        public const string IsdpLevel1 = "ISDP L1";
        public const string ExcelLevel1 = "Excel L1";
        public const string IsdpLevel2 = "ISDP L2";
        public const string ExcelLevel2 = "Excel L2";
        public const string HeaderId = "HeaderId";
        public const string ApiFieldKey = "ApiFieldKey";
        public const string IsIdColumn = "IsIdColumn";
        public const string ActivityId = "ActivityId";
        public const string PropertyId = "PropertyId";
    }

    public sealed class CurrentBusinessFieldMappingSeedBuilder
    {
        private readonly IReadOnlyDictionary<string, string> propertyLabels;

        public CurrentBusinessFieldMappingSeedBuilder(IReadOnlyDictionary<string, string> propertyLabels)
        {
            this.propertyLabels = propertyLabels ?? throw new ArgumentNullException(nameof(propertyLabels));
        }

        public IReadOnlyList<SheetFieldMappingRow> Build(
            string sheetName,
            IReadOnlyList<CurrentBusinessHeadDefinition> headList,
            IReadOnlyList<IDictionary<string, object>> sampleRows)
        {
            if (headList is null)
            {
                throw new ArgumentNullException(nameof(headList));
            }

            var sampleRowsToUse = sampleRows ?? Array.Empty<IDictionary<string, object>>();

            var mappings = new List<SheetFieldMappingRow>();

            foreach (var head in headList)
            {
                if (head == null)
                {
                    continue;
                }

                if (string.Equals(head.HeadType, "activity", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                mappings.Add(BuildSingleRow(sheetName, head));
            }

            var activityLookup = new Dictionary<string, CurrentBusinessHeadDefinition>(StringComparer.OrdinalIgnoreCase);

            foreach (var head in headList)
            {
                if (head == null || !string.Equals(head.HeadType, "activity", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(head.ActivityId))
                {
                    continue;
                }

                if (!activityLookup.TryGetValue(head.ActivityId, out var existing))
                {
                    activityLookup[head.ActivityId] = head;
                    continue;
                }

                if (string.IsNullOrWhiteSpace(existing?.ActivityName) && !string.IsNullOrWhiteSpace(head.ActivityName))
                {
                    activityLookup[head.ActivityId] = head;
                }
            }

            var propertyRows = new List<(string FlatKey, string PropertyKey, CurrentBusinessHeadDefinition Activity)>();

            foreach (var flatKey in sampleRowsToUse.Where(row => row != null).SelectMany(row => row.Keys).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                if (!TrySplitFlatKey(flatKey, out var propertyKey, out var activityId))
                {
                    continue;
                }

                if (!activityLookup.TryGetValue(activityId, out var activityHead))
                {
                    continue;
                }

                propertyRows.Add((flatKey, propertyKey, activityHead));
            }

            var orderedProperties = propertyRows
                .OrderBy(item => item.Activity.ActivityId, StringComparer.OrdinalIgnoreCase)
                .ThenBy(item => item.PropertyKey, StringComparer.OrdinalIgnoreCase);

            foreach (var property in orderedProperties)
            {
                mappings.Add(BuildActivityPropertyRow(sheetName, property.FlatKey, property.PropertyKey, property.Activity));
            }

            return mappings;
        }

        private SheetFieldMappingRow BuildSingleRow(string sheetName, CurrentBusinessHeadDefinition head)
        {
            var headerText = head?.HeaderText ?? string.Empty;
            var fieldKey = head?.FieldKey ?? string.Empty;
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                [CurrentBusinessFieldMappingColumns.HeaderId] = fieldKey,
                [CurrentBusinessFieldMappingColumns.HeaderType] = CurrentBusinessFieldMappingColumns.SingleHeaderType,
                [CurrentBusinessFieldMappingColumns.ApiFieldKey] = fieldKey,
                [CurrentBusinessFieldMappingColumns.IsIdColumn] = head?.IsId == true ? "true" : "false",
                [CurrentBusinessFieldMappingColumns.DefaultLevel1] = headerText,
                [CurrentBusinessFieldMappingColumns.CurrentLevel1] = headerText,
                [CurrentBusinessFieldMappingColumns.DefaultLevel2] = string.Empty,
                [CurrentBusinessFieldMappingColumns.CurrentLevel2] = string.Empty,
                [CurrentBusinessFieldMappingColumns.ActivityId] = string.Empty,
                [CurrentBusinessFieldMappingColumns.PropertyId] = string.Empty,
            };

            return new SheetFieldMappingRow
            {
                SheetName = sheetName ?? string.Empty,
                Values = values,
            };
        }

        private SheetFieldMappingRow BuildActivityPropertyRow(
            string sheetName,
            string flatKey,
            string propertyKey,
            CurrentBusinessHeadDefinition activity)
        {
            var activityName = activity?.ActivityName ?? string.Empty;
            var resolvedPropertyLabel = ResolvePropertyLabel(propertyKey);

            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                [CurrentBusinessFieldMappingColumns.HeaderId] = flatKey ?? string.Empty,
                [CurrentBusinessFieldMappingColumns.HeaderType] = CurrentBusinessFieldMappingColumns.ActivityPropertyHeaderType,
                [CurrentBusinessFieldMappingColumns.ApiFieldKey] = flatKey ?? string.Empty,
                [CurrentBusinessFieldMappingColumns.IsIdColumn] = "false",
                [CurrentBusinessFieldMappingColumns.DefaultLevel1] = activityName,
                [CurrentBusinessFieldMappingColumns.CurrentLevel1] = activityName,
                [CurrentBusinessFieldMappingColumns.DefaultLevel2] = resolvedPropertyLabel,
                [CurrentBusinessFieldMappingColumns.CurrentLevel2] = resolvedPropertyLabel,
                [CurrentBusinessFieldMappingColumns.ActivityId] = activity?.ActivityId ?? string.Empty,
                [CurrentBusinessFieldMappingColumns.PropertyId] = propertyKey ?? string.Empty,
            };

            return new SheetFieldMappingRow
            {
                SheetName = sheetName ?? string.Empty,
                Values = values,
            };
        }

        private string ResolvePropertyLabel(string propertyKey)
        {
            if (string.IsNullOrWhiteSpace(propertyKey))
            {
                return string.Empty;
            }

            return propertyLabels.TryGetValue(propertyKey, out var label)
                ? label ?? propertyKey
                : propertyKey;
        }

        private static bool TrySplitFlatKey(string flatKey, out string propertyKey, out string activityId)
        {
            propertyKey = null;
            activityId = null;

            if (string.IsNullOrWhiteSpace(flatKey))
            {
                return false;
            }

            var lastIndex = flatKey.LastIndexOf('_');
            if (lastIndex <= 0 || lastIndex == flatKey.Length - 1)
            {
                return false;
            }

            propertyKey = flatKey.Substring(0, lastIndex);
            activityId = flatKey.Substring(lastIndex + 1);
            return true;
        }
    }
}
