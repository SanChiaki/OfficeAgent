using System;
using System.Collections.Generic;
using System.Linq;
using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Sync
{
    public sealed class AiColumnMappingService
    {
        private const double ConfidenceThreshold = 0.75;

        private readonly FieldMappingValueAccessor accessor = new FieldMappingValueAccessor();

        public AiColumnMappingRequest BuildRequest(
            string sheetName,
            FieldMappingTableDefinition definition,
            IReadOnlyList<SheetFieldMappingRow> rows,
            IReadOnlyList<AiColumnMappingActualHeader> actualHeaders)
        {
            if (definition == null)
            {
                throw new ArgumentNullException(nameof(definition));
            }

            return new AiColumnMappingRequest
            {
                SystemKey = definition.SystemKey ?? string.Empty,
                SheetName = sheetName ?? string.Empty,
                ActualHeaders = (actualHeaders ?? Array.Empty<AiColumnMappingActualHeader>())
                    .Where(header => header != null)
                    .ToArray(),
                Candidates = (rows ?? Array.Empty<SheetFieldMappingRow>())
                    .Where(row => IsTargetSheet(row, sheetName))
                    .Select(row => CreateCandidate(definition, row))
                    .ToArray(),
            };
        }

        public AiColumnMappingPreview CreatePreview(
            AiColumnMappingRequest request,
            AiColumnMappingResponse response,
            int headerRowCount)
        {
            var candidates = (request?.Candidates ?? Array.Empty<AiColumnMappingCandidate>())
                .Where(candidate => candidate != null)
                .ToArray();
            var actualHeaders = (request?.ActualHeaders ?? Array.Empty<AiColumnMappingActualHeader>())
                .Where(header => header != null)
                .ToArray();
            var candidatesByHeaderId = candidates
                .Where(candidate => !string.IsNullOrWhiteSpace(candidate.HeaderId))
                .GroupBy(candidate => candidate.HeaderId, StringComparer.OrdinalIgnoreCase)
                .Where(group => group.Count() == 1)
                .ToDictionary(group => group.Key, group => group.Single(), StringComparer.OrdinalIgnoreCase);
            var candidatesByApiFieldKey = candidates
                .Where(candidate => !string.IsNullOrWhiteSpace(candidate.ApiFieldKey))
                .GroupBy(candidate => candidate.ApiFieldKey, StringComparer.OrdinalIgnoreCase)
                .Where(group => group.Count() == 1)
                .ToDictionary(group => group.Key, group => group.Single(), StringComparer.OrdinalIgnoreCase);
            var actualByColumn = actualHeaders
                .GroupBy(header => header.ExcelColumn)
                .ToDictionary(group => group.Key, group => group.First());
            var items = new List<AiColumnMappingPreviewItem>();
            var suggestions = new List<PreviewSuggestion>();

            foreach (var suggestion in response?.Mappings ?? Array.Empty<AiColumnMappingSuggestion>())
            {
                if (suggestion == null)
                {
                    continue;
                }

                var hasActualHeader = actualByColumn.TryGetValue(suggestion.ExcelColumn, out var actual);
                var candidate = ResolveCandidate(candidatesByHeaderId, candidatesByApiFieldKey, suggestion);
                var item = CreatePreviewItem(actual, candidate, suggestion);
                suggestions.Add(new PreviewSuggestion(item, candidate, hasActualHeader));
            }

            var duplicateTargets = CreateDuplicateTargetIdentities(
                suggestions
                    .Where(suggestion => suggestion.HasActualHeader)
                    .Select(suggestion => suggestion.Candidate));
            var duplicateColumns = new HashSet<int>(
                suggestions
                    .GroupBy(suggestion => suggestion.Item.ExcelColumn)
                    .Where(group => group.Count() > 1)
                    .Select(group => group.Key));

            foreach (var suggestion in suggestions)
            {
                suggestion.Item.Status = ResolveStatus(suggestion.Item, suggestion.Candidate, suggestion.HasActualHeader, headerRowCount, duplicateTargets, duplicateColumns);
                suggestion.Item.Reason = ResolveReason(suggestion.Item.Status, suggestion.Item, suggestion.Candidate, suggestion.HasActualHeader, headerRowCount, duplicateTargets, duplicateColumns);
                items.Add(suggestion.Item);
            }

            foreach (var unmatched in response?.Unmatched ?? Array.Empty<AiColumnMappingUnmatchedHeader>())
            {
                if (unmatched == null)
                {
                    continue;
                }

                items.Add(new AiColumnMappingPreviewItem
                {
                    ExcelColumn = unmatched.ExcelColumn,
                    SuggestedExcelL1 = FirstNonEmpty(unmatched.ActualL1, unmatched.DisplayText),
                    SuggestedExcelL2 = unmatched.ActualL2 ?? string.Empty,
                    Status = AiColumnMappingPreviewStatuses.Unmatched,
                    Reason = unmatched.Reason ?? string.Empty,
                });
            }

            return new AiColumnMappingPreview
            {
                Items = items.ToArray(),
            };
        }

        public AiColumnMappingApplyResult ApplyConfirmedPreview(
            string sheetName,
            FieldMappingTableDefinition definition,
            IReadOnlyList<SheetFieldMappingRow> rows,
            AiColumnMappingPreview preview,
            int headerRowCount)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                throw new ArgumentException("Sheet name is required.", nameof(sheetName));
            }

            if (definition == null)
            {
                throw new ArgumentNullException(nameof(definition));
            }

            var acceptedPreviewItems = (preview?.Items ?? Array.Empty<AiColumnMappingPreviewItem>())
                .Where(item => item != null && string.Equals(item.Status, AiColumnMappingPreviewStatuses.Accepted, StringComparison.Ordinal))
                .ToArray();
            var duplicateColumns = new HashSet<int>(
                acceptedPreviewItems
                    .GroupBy(item => item.ExcelColumn)
                    .Where(group => group.Count() > 1)
                    .Select(group => group.Key));
            var duplicateTargets = new HashSet<string>(
                acceptedPreviewItems
                    .Where(item => !string.IsNullOrWhiteSpace(CreateTargetKey(item)))
                    .GroupBy(CreateTargetKey, StringComparer.OrdinalIgnoreCase)
                    .Where(group => group.Count() > 1)
                    .Select(group => group.Key),
                StringComparer.OrdinalIgnoreCase);
            foreach (var duplicateTarget in CreateDuplicateTargetIdentities(acceptedPreviewItems))
            {
                duplicateTargets.Add(duplicateTarget);
            }

            var acceptedByTarget = acceptedPreviewItems
                .Where(item => CanApply(item, headerRowCount))
                .Where(item => !duplicateColumns.Contains(item.ExcelColumn))
                .Where(item => !HasDuplicateTargetIdentity(duplicateTargets, item))
                .GroupBy(CreateTargetKey, StringComparer.OrdinalIgnoreCase)
                .Where(group => group.Count() == 1)
                .ToDictionary(group => group.Key, group => group.Single(), StringComparer.OrdinalIgnoreCase);
            var clonedRows = (rows ?? Array.Empty<SheetFieldMappingRow>())
                .Where(row => row != null)
                .Select(CloneRow)
                .ToArray();
            var appliedCount = 0;

            foreach (var row in clonedRows)
            {
                if (!IsTargetSheet(row, sheetName))
                {
                    continue;
                }

                var key = CreateTargetKey(
                    accessor.GetValue(definition, row, FieldMappingSemanticRole.HeaderIdentity),
                    accessor.GetValue(definition, row, FieldMappingSemanticRole.ApiFieldKey));
                if (acceptedByTarget.TryGetValue(key, out var accepted) && MatchesTarget(definition, row, accepted))
                {
                    ApplyAccepted(definition, row, accepted);
                    appliedCount++;
                }
            }

            return new AiColumnMappingApplyResult
            {
                Rows = clonedRows,
                AppliedCount = appliedCount,
                SkippedCount = Math.Max(0, (preview?.Items ?? Array.Empty<AiColumnMappingPreviewItem>()).Length - appliedCount),
            };
        }

        private AiColumnMappingCandidate CreateCandidate(
            FieldMappingTableDefinition definition,
            SheetFieldMappingRow row)
        {
            var defaultSingle = accessor.GetValue(definition, row, FieldMappingSemanticRole.DefaultSingleHeaderText);
            var defaultParent = accessor.GetValue(definition, row, FieldMappingSemanticRole.DefaultParentHeaderText);

            return new AiColumnMappingCandidate
            {
                HeaderId = accessor.GetValue(definition, row, FieldMappingSemanticRole.HeaderIdentity),
                ApiFieldKey = accessor.GetValue(definition, row, FieldMappingSemanticRole.ApiFieldKey),
                HeaderType = accessor.GetValue(definition, row, FieldMappingSemanticRole.HeaderType),
                IsdpL1 = string.IsNullOrWhiteSpace(defaultSingle) ? defaultParent : defaultSingle,
                IsdpL2 = accessor.GetValue(definition, row, FieldMappingSemanticRole.DefaultChildHeaderText),
                CurrentExcelL1 = ResolveCurrentL1(definition, row),
                CurrentExcelL2 = accessor.GetValue(definition, row, FieldMappingSemanticRole.CurrentChildHeaderText),
                IsIdColumn = accessor.GetBoolean(definition, row, FieldMappingSemanticRole.IsIdColumn),
                ActivityId = accessor.GetValue(definition, row, FieldMappingSemanticRole.ActivityIdentity),
                PropertyId = accessor.GetValue(definition, row, FieldMappingSemanticRole.PropertyIdentity),
            };
        }

        private string ResolveCurrentL1(
            FieldMappingTableDefinition definition,
            SheetFieldMappingRow row)
        {
            var currentSingle = accessor.GetValue(definition, row, FieldMappingSemanticRole.CurrentSingleHeaderText);
            var currentParent = accessor.GetValue(definition, row, FieldMappingSemanticRole.CurrentParentHeaderText);
            return string.IsNullOrWhiteSpace(currentSingle) ? currentParent : currentSingle;
        }

        private static AiColumnMappingCandidate ResolveCandidate(
            IReadOnlyDictionary<string, AiColumnMappingCandidate> candidatesByHeaderId,
            IReadOnlyDictionary<string, AiColumnMappingCandidate> candidatesByApiFieldKey,
            AiColumnMappingSuggestion suggestion)
        {
            AiColumnMappingCandidate byHeaderId = null;
            AiColumnMappingCandidate byApiFieldKey = null;

            var hasHeaderId = !string.IsNullOrWhiteSpace(suggestion.TargetHeaderId);
            var hasApiFieldKey = !string.IsNullOrWhiteSpace(suggestion.TargetApiFieldKey);

            if (hasHeaderId)
            {
                candidatesByHeaderId?.TryGetValue(suggestion.TargetHeaderId, out byHeaderId);
                if (byHeaderId == null)
                {
                    return null;
                }
            }

            if (hasApiFieldKey)
            {
                candidatesByApiFieldKey?.TryGetValue(suggestion.TargetApiFieldKey, out byApiFieldKey);
                if (byApiFieldKey == null)
                {
                    return null;
                }
            }

            if (byHeaderId != null && byApiFieldKey != null && !ReferenceEquals(byHeaderId, byApiFieldKey))
            {
                return null;
            }

            return byHeaderId ?? byApiFieldKey;
        }

        private static AiColumnMappingPreviewItem CreatePreviewItem(
            AiColumnMappingActualHeader actual,
            AiColumnMappingCandidate candidate,
            AiColumnMappingSuggestion suggestion)
        {
            return new AiColumnMappingPreviewItem
            {
                ExcelColumn = suggestion.ExcelColumn,
                SuggestedExcelL1 = FirstNonEmpty(actual?.ActualL1, FirstNonEmpty(suggestion.ActualL1, actual?.DisplayText)),
                SuggestedExcelL2 = FirstNonEmpty(actual?.ActualL2, suggestion.ActualL2),
                TargetHeaderId = candidate?.HeaderId ?? suggestion.TargetHeaderId ?? string.Empty,
                TargetApiFieldKey = candidate?.ApiFieldKey ?? suggestion.TargetApiFieldKey ?? string.Empty,
                HeaderType = candidate?.HeaderType ?? string.Empty,
                TargetIsdpL1 = candidate?.DefaultL1 ?? string.Empty,
                TargetIsdpL2 = candidate?.DefaultL2 ?? string.Empty,
                Confidence = suggestion.Confidence,
                Reason = suggestion.Reason ?? string.Empty,
            };
        }

        private static string ResolveStatus(
            AiColumnMappingPreviewItem item,
            AiColumnMappingCandidate candidate,
            bool hasActualHeader,
            int headerRowCount,
            ISet<string> duplicateTargets,
            ISet<int> duplicateColumns)
        {
            if (!hasActualHeader)
            {
                return AiColumnMappingPreviewStatuses.Rejected;
            }

            if (candidate == null || string.IsNullOrWhiteSpace(CreateTargetKey(candidate)))
            {
                return AiColumnMappingPreviewStatuses.Rejected;
            }

            if (HasDuplicateTargetIdentity(duplicateTargets, candidate))
            {
                return AiColumnMappingPreviewStatuses.Rejected;
            }

            if (duplicateColumns.Contains(item.ExcelColumn))
            {
                return AiColumnMappingPreviewStatuses.Rejected;
            }

            return item.Confidence >= ConfidenceThreshold
                ? AiColumnMappingPreviewStatuses.Accepted
                : AiColumnMappingPreviewStatuses.LowConfidence;
        }

        private static string ResolveReason(
            string status,
            AiColumnMappingPreviewItem item,
            AiColumnMappingCandidate candidate,
            bool hasActualHeader,
            int headerRowCount,
            ISet<string> duplicateTargets,
            ISet<int> duplicateColumns)
        {
            if (!string.Equals(status, AiColumnMappingPreviewStatuses.Rejected, StringComparison.Ordinal))
            {
                return string.Empty;
            }

            if (candidate == null || string.IsNullOrWhiteSpace(CreateTargetKey(candidate)))
            {
                return "Invalid target identity.";
            }

            if (!hasActualHeader)
            {
                return "Rejected missing actual header.";
            }

            if (HasDuplicateTargetIdentity(duplicateTargets, candidate))
            {
                return "Rejected duplicate target field.";
            }

            if (duplicateColumns.Contains(item.ExcelColumn))
            {
                return "Rejected duplicate Excel column.";
            }

            return "Rejected suggestion.";
        }

        private static bool CanApply(AiColumnMappingPreviewItem item, int headerRowCount)
        {
            return item != null
                && item.ExcelColumn > 0
                && !string.IsNullOrWhiteSpace(item.TargetHeaderId)
                && !string.IsNullOrWhiteSpace(item.TargetApiFieldKey)
                && item.Confidence >= ConfidenceThreshold;
        }

        private static HashSet<string> CreateDuplicateTargetIdentities(IEnumerable<AiColumnMappingCandidate> candidates)
        {
            return CreateDuplicateTargetIdentities(
                candidates,
                candidate => candidate?.HeaderId,
                candidate => candidate?.ApiFieldKey);
        }

        private static HashSet<string> CreateDuplicateTargetIdentities(IEnumerable<AiColumnMappingPreviewItem> items)
        {
            return CreateDuplicateTargetIdentities(
                items,
                item => item?.TargetHeaderId,
                item => item?.TargetApiFieldKey);
        }

        private static HashSet<string> CreateDuplicateTargetIdentities<T>(
            IEnumerable<T> values,
            Func<T, string> getHeaderId,
            Func<T, string> getApiFieldKey)
        {
            var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            AddDuplicateValues(result, values, getHeaderId);
            AddDuplicateValues(result, values, getApiFieldKey);
            return result;
        }

        private static void AddDuplicateValues<T>(
            ISet<string> result,
            IEnumerable<T> values,
            Func<T, string> getValue)
        {
            foreach (var duplicateValue in (values ?? Enumerable.Empty<T>())
                .Select(getValue)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .GroupBy(value => value, StringComparer.OrdinalIgnoreCase)
                .Where(group => group.Count() > 1)
                .Select(group => group.Key))
            {
                result.Add(duplicateValue);
            }
        }

        private static bool HasDuplicateTargetIdentity(
            ISet<string> duplicateTargets,
            AiColumnMappingCandidate candidate)
        {
            return candidate != null &&
                (ContainsDuplicateTarget(duplicateTargets, candidate.HeaderId) ||
                 ContainsDuplicateTarget(duplicateTargets, candidate.ApiFieldKey));
        }

        private static bool HasDuplicateTargetIdentity(
            ISet<string> duplicateTargets,
            AiColumnMappingPreviewItem item)
        {
            return item != null &&
                (ContainsDuplicateTarget(duplicateTargets, item.TargetHeaderId) ||
                 ContainsDuplicateTarget(duplicateTargets, item.TargetApiFieldKey));
        }

        private static bool ContainsDuplicateTarget(ISet<string> duplicateTargets, string value)
        {
            return !string.IsNullOrWhiteSpace(value) &&
                   duplicateTargets != null &&
                   duplicateTargets.Contains(value);
        }

        private bool MatchesTarget(
            FieldMappingTableDefinition definition,
            SheetFieldMappingRow row,
            AiColumnMappingPreviewItem item)
        {
            var rowHeaderId = accessor.GetValue(definition, row, FieldMappingSemanticRole.HeaderIdentity);
            var rowApiFieldKey = accessor.GetValue(definition, row, FieldMappingSemanticRole.ApiFieldKey);

            return !string.IsNullOrWhiteSpace(item.TargetHeaderId)
                && !string.IsNullOrWhiteSpace(item.TargetApiFieldKey)
                && string.Equals(item.TargetHeaderId, rowHeaderId, StringComparison.OrdinalIgnoreCase)
                && string.Equals(item.TargetApiFieldKey, rowApiFieldKey, StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsTargetSheet(SheetFieldMappingRow row, string sheetName)
        {
            return row != null
                && !string.IsNullOrWhiteSpace(sheetName)
                && string.Equals(row.SheetName, sheetName, StringComparison.OrdinalIgnoreCase);
        }

        private static SheetFieldMappingRow CloneRow(SheetFieldMappingRow row)
        {
            return new SheetFieldMappingRow
            {
                SheetName = row.SheetName,
                Values = CopyValues(row.Values),
            };
        }

        private static void ApplyAccepted(
            FieldMappingTableDefinition definition,
            SheetFieldMappingRow row,
            AiColumnMappingPreviewItem accepted)
        {
            var values = row.Values as IDictionary<string, string>;
            if (values == null)
            {
                values = CopyValues(row.Values);
                row.Values = (IReadOnlyDictionary<string, string>)values;
            }

            SetValue(definition, values, FieldMappingSemanticRole.CurrentSingleHeaderText, accepted.ActualL1);
            SetValue(definition, values, FieldMappingSemanticRole.CurrentParentHeaderText, accepted.ActualL1);
            SetValue(definition, values, FieldMappingSemanticRole.CurrentChildHeaderText, accepted.ActualL2);
        }

        private static void SetValue(
            FieldMappingTableDefinition definition,
            IDictionary<string, string> values,
            FieldMappingSemanticRole role,
            string value)
        {
            var key = (definition.Columns ?? Array.Empty<FieldMappingColumnDefinition>())
                .Where(column => column != null && column.Role == role)
                .Select(ResolveValueKey)
                .FirstOrDefault(name => !string.IsNullOrWhiteSpace(name));
            if (!string.IsNullOrWhiteSpace(key))
            {
                values[key] = value ?? string.Empty;
            }
        }

        private static Dictionary<string, string> CopyValues(IReadOnlyDictionary<string, string> values)
        {
            var copy = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (values == null)
            {
                return copy;
            }

            foreach (var pair in values)
            {
                copy[pair.Key] = pair.Value;
            }

            return copy;
        }

        private static string ResolveValueKey(FieldMappingColumnDefinition column)
        {
            return string.IsNullOrWhiteSpace(column.RoleKey)
                ? column.ColumnName ?? string.Empty
                : column.RoleKey;
        }

        private static string CreateTargetKey(AiColumnMappingCandidate candidate)
        {
            return candidate == null
                ? string.Empty
                : CreateTargetKey(candidate.HeaderId, candidate.ApiFieldKey);
        }

        private static string CreateTargetKey(AiColumnMappingPreviewItem item)
        {
            return item == null
                ? string.Empty
                : CreateTargetKey(item.HeaderId, item.ApiFieldKey);
        }

        private static string CreateTargetKey(string headerId, string apiFieldKey)
        {
            var identity = !string.IsNullOrWhiteSpace(headerId) ? headerId : apiFieldKey;
            return identity ?? string.Empty;
        }

        private static string FirstNonEmpty(string first, string second)
        {
            return !string.IsNullOrWhiteSpace(first) ? first : second ?? string.Empty;
        }

        private sealed class PreviewSuggestion
        {
            public PreviewSuggestion(AiColumnMappingPreviewItem item, AiColumnMappingCandidate candidate, bool hasActualHeader)
            {
                Item = item;
                Candidate = candidate;
                HasActualHeader = hasActualHeader;
            }

            public AiColumnMappingPreviewItem Item { get; }

            public AiColumnMappingCandidate Candidate { get; }

            public bool HasActualHeader { get; }
        }
    }
}
