using System;

namespace OfficeAgent.ExcelAddIn.Updates
{
    internal sealed class UpdateState
    {
        public DateTime? LastCheckedAtUtc { get; set; }
        public string LatestVersion { get; set; } = string.Empty;
        public string DownloadUrl { get; set; } = string.Empty;
        public string ReleaseNotesUrl { get; set; } = string.Empty;
        public DateTime? PublishedAtUtc { get; set; }
        public string Title { get; set; } = string.Empty;
        public string Summary { get; set; } = string.Empty;
        public string IgnoredVersion { get; set; } = string.Empty;

        public UpdateManifest ToManifest()
        {
            return new UpdateManifest
            {
                LatestVersion = LatestVersion ?? string.Empty,
                DownloadUrl = DownloadUrl ?? string.Empty,
                ReleaseNotesUrl = ReleaseNotesUrl ?? string.Empty,
                PublishedAtUtc = PublishedAtUtc,
                Title = Title ?? string.Empty,
                Summary = Summary ?? string.Empty,
            };
        }

        public void ApplyManifest(UpdateManifest manifest, DateTime checkedAtUtc)
        {
            LastCheckedAtUtc = checkedAtUtc;
            LatestVersion = manifest?.LatestVersion ?? string.Empty;
            DownloadUrl = manifest?.DownloadUrl ?? string.Empty;
            ReleaseNotesUrl = manifest?.ReleaseNotesUrl ?? string.Empty;
            PublishedAtUtc = manifest?.PublishedAtUtc;
            Title = manifest?.Title ?? string.Empty;
            Summary = manifest?.Summary ?? string.Empty;
        }
    }
}
