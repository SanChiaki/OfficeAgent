using System;

namespace OfficeAgent.ExcelAddIn.Updates
{
    internal sealed class UpdateManifest
    {
        public string LatestVersion { get; set; } = string.Empty;
        public string DownloadUrl { get; set; } = string.Empty;
        public string ReleaseNotesUrl { get; set; } = string.Empty;
        public DateTime? PublishedAtUtc { get; set; }
        public string Title { get; set; } = string.Empty;
        public string Summary { get; set; } = string.Empty;
    }
}
