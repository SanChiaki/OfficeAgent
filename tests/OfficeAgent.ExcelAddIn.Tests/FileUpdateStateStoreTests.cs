using System;
using System.IO;
using OfficeAgent.ExcelAddIn.Updates;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class FileUpdateStateStoreTests
    {
        [Fact]
        public void SaveAndLoadRoundTripsUpdateState()
        {
            var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"), "state.json");
            var state = new UpdateState
            {
                LastCheckedAtUtc = new DateTime(2026, 5, 19, 8, 30, 0, DateTimeKind.Utc),
                LatestVersion = "1.0.176",
                DownloadUrl = "https://updates.example/download.exe",
                ReleaseNotesUrl = "https://updates.example/notes",
                PublishedAtUtc = new DateTime(2026, 5, 19, 8, 0, 0, DateTimeKind.Utc),
                Title = "Release",
                Summary = "Summary",
                IgnoredVersion = "1.0.175",
            };

            var store = new FileUpdateStateStore(path);

            store.Save(state);
            var loaded = store.Load();

            Assert.Equal(state.LastCheckedAtUtc, loaded.LastCheckedAtUtc);
            Assert.Equal(state.LatestVersion, loaded.LatestVersion);
            Assert.Equal(state.DownloadUrl, loaded.DownloadUrl);
            Assert.Equal(state.ReleaseNotesUrl, loaded.ReleaseNotesUrl);
            Assert.Equal(state.PublishedAtUtc, loaded.PublishedAtUtc);
            Assert.Equal(state.Title, loaded.Title);
            Assert.Equal(state.Summary, loaded.Summary);
            Assert.Equal(state.IgnoredVersion, loaded.IgnoredVersion);
        }

        [Fact]
        public void LoadReturnsEmptyStateWhenFileIsMissingOrCorrupt()
        {
            var missingPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"), "missing.json");
            var missingState = new FileUpdateStateStore(missingPath).Load();

            Assert.Equal(string.Empty, missingState.LatestVersion);

            var corruptPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"), "state.json");
            Directory.CreateDirectory(Path.GetDirectoryName(corruptPath));
            File.WriteAllText(corruptPath, "not-json");

            var corruptState = new FileUpdateStateStore(corruptPath).Load();

            Assert.Equal(string.Empty, corruptState.LatestVersion);
        }
    }
}
