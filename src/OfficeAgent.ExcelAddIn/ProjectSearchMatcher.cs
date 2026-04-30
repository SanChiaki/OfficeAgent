using System;
using System.Linq;

namespace OfficeAgent.ExcelAddIn
{
    internal static class ProjectSearchMatcher
    {
        public static bool IsMatch(string label, string query)
        {
            var normalizedLabel = label ?? string.Empty;
            var normalizedQuery = query?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(normalizedQuery))
            {
                return true;
            }

            if (normalizedLabel.IndexOf(normalizedQuery, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            var terms = normalizedQuery
                .Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (terms.Length > 1 &&
                terms.All(term => normalizedLabel.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0))
            {
                return true;
            }

            return IsSubsequence(Compact(normalizedLabel), Compact(normalizedQuery));
        }

        private static string Compact(string value)
        {
            return new string((value ?? string.Empty)
                .Where(c => !char.IsWhiteSpace(c) && c != '-' && c != '_' && c != '.')
                .Select(char.ToLowerInvariant)
                .ToArray());
        }

        private static bool IsSubsequence(string label, string query)
        {
            if (string.IsNullOrWhiteSpace(query))
            {
                return true;
            }

            var queryIndex = 0;
            foreach (var c in label ?? string.Empty)
            {
                if (queryIndex < query.Length && c == query[queryIndex])
                {
                    queryIndex++;
                    if (queryIndex == query.Length)
                    {
                        return true;
                    }
                }
            }

            return false;
        }
    }
}
