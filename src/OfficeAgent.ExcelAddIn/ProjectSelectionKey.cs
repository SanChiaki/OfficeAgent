using System;

namespace OfficeAgent.ExcelAddIn
{
    internal static class ProjectSelectionKey
    {
        private const string Separator = "::";

        public static string Build(string systemKey, string projectId)
        {
            if (string.IsNullOrWhiteSpace(systemKey))
            {
                throw new ArgumentException("System key is required.", nameof(systemKey));
            }

            if (string.IsNullOrWhiteSpace(projectId))
            {
                throw new ArgumentException("Project id is required.", nameof(projectId));
            }

            return string.Concat(
                Uri.EscapeDataString(systemKey.Trim()),
                Separator,
                Uri.EscapeDataString(projectId.Trim()));
        }

        public static bool TryParse(string value, out string systemKey, out string projectId)
        {
            systemKey = null;
            projectId = null;

            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            var separatorIndex = value.IndexOf(Separator, StringComparison.Ordinal);
            if (separatorIndex <= 0 || separatorIndex >= value.Length - Separator.Length)
            {
                return false;
            }

            var encodedSystemKey = value.Substring(0, separatorIndex);
            var encodedProjectId = value.Substring(separatorIndex + Separator.Length);
            if (string.IsNullOrWhiteSpace(encodedSystemKey) || string.IsNullOrWhiteSpace(encodedProjectId))
            {
                return false;
            }

            systemKey = Uri.UnescapeDataString(encodedSystemKey);
            projectId = Uri.UnescapeDataString(encodedProjectId);
            return !string.IsNullOrWhiteSpace(systemKey) && !string.IsNullOrWhiteSpace(projectId);
        }
    }
}
