using System;

namespace OfficeAgent.ExcelAddIn.Updates
{
    internal static class UpdateVersionComparer
    {
        public static bool IsNewerThanCurrent(string latestVersion, string currentVersion)
        {
            return TryParse(latestVersion, out var latest) &&
                   TryParse(currentVersion, out var current) &&
                   latest.CompareTo(current) > 0;
        }

        public static bool TryParse(string value, out Version version)
        {
            version = null;
            var normalized = (value ?? string.Empty).Trim();
            if (normalized.StartsWith("v", StringComparison.OrdinalIgnoreCase))
            {
                normalized = normalized.Substring(1);
            }

            return Version.TryParse(normalized, out version);
        }
    }
}
