using System;

namespace OfficeAgent.ExcelAddIn.Updates
{
    internal sealed class UpdateNotificationState
    {
        public static readonly UpdateNotificationState Empty = new UpdateNotificationState(false, false, string.Empty, string.Empty, string.Empty, null, string.Empty, string.Empty);

        public UpdateNotificationState(bool hasNewVersion, bool shouldShowReminder, string latestVersion, string downloadUrl, string releaseNotesUrl, DateTime? publishedAtUtc, string title, string summary)
        {
            HasNewVersion = hasNewVersion;
            ShouldShowReminder = hasNewVersion && shouldShowReminder;
            LatestVersion = latestVersion ?? string.Empty;
            DownloadUrl = downloadUrl ?? string.Empty;
            ReleaseNotesUrl = releaseNotesUrl ?? string.Empty;
            PublishedAtUtc = publishedAtUtc;
            Title = title ?? string.Empty;
            Summary = summary ?? string.Empty;
        }

        public bool HasNewVersion { get; }
        public bool ShouldShowReminder { get; }
        public string LatestVersion { get; }
        public string DownloadUrl { get; }
        public string ReleaseNotesUrl { get; }
        public DateTime? PublishedAtUtc { get; }
        public string Title { get; }
        public string Summary { get; }
    }
}
