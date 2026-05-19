using System;
using System.Threading;
using System.Threading.Tasks;
using OfficeAgent.ExcelAddIn.Updates;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class UpdateNotificationServiceTests
    {
        private static readonly DateTime NowUtc = new DateTime(2026, 5, 19, 8, 0, 0, DateTimeKind.Utc);

        [Fact]
        public async Task CheckForUpdatesAsyncSkipsHttpWhenDisabled()
        {
            var client = new FakeUpdateManifestClient();
            var store = new MemoryUpdateStateStore();
            var service = CreateService(UpdateCheckOptions.Disabled(), client, store);

            await service.CheckForUpdatesAsync(CancellationToken.None);

            Assert.Equal(0, client.CallCount);
            Assert.False(service.CurrentState.HasNewVersion);
        }

        [Fact]
        public async Task CheckForUpdatesAsyncUsesCacheWindowWithoutHttp()
        {
            var client = new FakeUpdateManifestClient();
            var store = new MemoryUpdateStateStore
            {
                State = new UpdateState
                {
                    LastCheckedAtUtc = NowUtc.AddHours(-1),
                    LatestVersion = "1.0.176",
                },
            };
            var service = CreateService(UpdateCheckOptions.Enabled("https://updates.example/manifest.json"), client, store);

            await service.CheckForUpdatesAsync(CancellationToken.None);

            Assert.Equal(0, client.CallCount);
            Assert.True(service.CurrentState.HasNewVersion);
            Assert.Equal("1.0.176", service.CurrentState.LatestVersion);
        }

        [Fact]
        public async Task CheckForUpdatesAsyncShowsNewVersionWhenLatestIsHigherAndNotIgnored()
        {
            var publishedAtUtc = new DateTime(2026, 5, 18, 8, 0, 0, DateTimeKind.Utc);
            var client = new FakeUpdateManifestClient
            {
                Manifest = new UpdateManifest
                {
                    LatestVersion = "1.0.176",
                    DownloadUrl = "https://updates.example/download.exe",
                    ReleaseNotesUrl = "https://updates.example/notes",
                    PublishedAtUtc = publishedAtUtc,
                    Title = "Release",
                    Summary = "Summary",
                },
            };
            var store = new MemoryUpdateStateStore();
            var service = CreateService(UpdateCheckOptions.Enabled("https://updates.example/manifest.json"), client, store);

            await service.CheckForUpdatesAsync(CancellationToken.None);

            Assert.Equal(1, client.CallCount);
            Assert.True(service.CurrentState.HasNewVersion);
            Assert.Equal("1.0.176", service.CurrentState.LatestVersion);
            Assert.Equal("https://updates.example/download.exe", service.CurrentState.DownloadUrl);
            Assert.Equal("https://updates.example/notes", service.CurrentState.ReleaseNotesUrl);
            Assert.Equal(publishedAtUtc, service.CurrentState.PublishedAtUtc);
            Assert.Equal("Release", service.CurrentState.Title);
            Assert.Equal("Summary", service.CurrentState.Summary);
            Assert.Equal(NowUtc, store.State.LastCheckedAtUtc);
            Assert.Equal("1.0.176", store.State.LatestVersion);
            Assert.Equal("https://updates.example/download.exe", store.State.DownloadUrl);
            Assert.Equal("https://updates.example/notes", store.State.ReleaseNotesUrl);
            Assert.Equal(publishedAtUtc, store.State.PublishedAtUtc);
            Assert.Equal("Release", store.State.Title);
            Assert.Equal("Summary", store.State.Summary);
        }

        [Fact]
        public async Task CheckForUpdatesAsyncHidesIgnoredVersionButShowsHigherFutureVersion()
        {
            var client = new FakeUpdateManifestClient
            {
                Manifest = new UpdateManifest
                {
                    LatestVersion = "1.0.176",
                },
            };
            var store = new MemoryUpdateStateStore
            {
                State = new UpdateState
                {
                    IgnoredVersion = "1.0.176",
                },
            };
            var service = CreateService(UpdateCheckOptions.Enabled("https://updates.example/manifest.json"), client, store);

            await service.CheckForUpdatesAsync(CancellationToken.None);

            Assert.False(service.CurrentState.HasNewVersion);

            client.Manifest = new UpdateManifest
            {
                LatestVersion = "1.0.177",
            };
            store.State.LastCheckedAtUtc = NowUtc.AddDays(-2);

            await service.CheckForUpdatesAsync(CancellationToken.None);

            Assert.True(service.CurrentState.HasNewVersion);
            Assert.Equal("1.0.177", service.CurrentState.LatestVersion);
            Assert.Equal(2, client.CallCount);
        }

        [Fact]
        public async Task IgnoreCurrentVersionPersistsIgnoredVersionAndRaisesStateChanged()
        {
            var client = new FakeUpdateManifestClient
            {
                Manifest = new UpdateManifest
                {
                    LatestVersion = "1.0.176",
                },
            };
            var store = new MemoryUpdateStateStore();
            var service = CreateService(UpdateCheckOptions.Enabled("https://updates.example/manifest.json"), client, store);
            var stateChangedCount = 0;
            service.StateChanged += (sender, args) => stateChangedCount++;
            await service.CheckForUpdatesAsync(CancellationToken.None);

            service.IgnoreCurrentVersion();

            Assert.Equal("1.0.176", store.State.IgnoredVersion);
            Assert.False(service.CurrentState.HasNewVersion);
            Assert.True(stateChangedCount >= 1);
        }

        [Fact]
        public void StartBackgroundCheckReturnsBeforeTheDelayedRequestRuns()
        {
            var client = new FakeUpdateManifestClient();
            var store = new MemoryUpdateStateStore();
            var service = CreateService(
                UpdateCheckOptions.Enabled(
                    "https://updates.example/manifest.json",
                    startupDelay: TimeSpan.FromMilliseconds(200)),
                client,
                store);

            service.StartBackgroundCheck(null);

            Assert.Equal(0, client.CallCount);
        }

        private static UpdateNotificationService CreateService(UpdateCheckOptions options, FakeUpdateManifestClient client, MemoryUpdateStateStore store)
        {
            return new UpdateNotificationService(options, client, store, "1.0.175", () => NowUtc);
        }

        private sealed class FakeUpdateManifestClient : IUpdateManifestClient
        {
            public UpdateManifest Manifest { get; set; } = new UpdateManifest
            {
                LatestVersion = "1.0.176",
            };

            public int CallCount { get; private set; }

            public Task<UpdateManifest> GetManifestAsync(string manifestUrl, CancellationToken cancellationToken)
            {
                CallCount++;
                return Task.FromResult(Manifest);
            }
        }

        private sealed class MemoryUpdateStateStore : IUpdateStateStore
        {
            public UpdateState State { get; set; } = new UpdateState();

            public UpdateState Load()
            {
                return Clone(State);
            }

            public void Save(UpdateState state)
            {
                State = Clone(state);
            }

            private static UpdateState Clone(UpdateState state)
            {
                return new UpdateState
                {
                    LastCheckedAtUtc = state.LastCheckedAtUtc,
                    LatestVersion = state.LatestVersion,
                    DownloadUrl = state.DownloadUrl,
                    ReleaseNotesUrl = state.ReleaseNotesUrl,
                    PublishedAtUtc = state.PublishedAtUtc,
                    Title = state.Title,
                    Summary = state.Summary,
                    IgnoredVersion = state.IgnoredVersion,
                };
            }
        }
    }
}
