using System;
using System.Threading;
using System.Threading.Tasks;
using OfficeAgent.Core.Diagnostics;

namespace OfficeAgent.ExcelAddIn.Updates
{
    internal sealed class UpdateNotificationService
    {
        private readonly object syncRoot = new object();
        private readonly UpdateCheckOptions options;
        private readonly IUpdateManifestClient manifestClient;
        private readonly IUpdateStateStore stateStore;
        private readonly string currentVersion;
        private readonly Func<DateTime> getUtcNow;
        private UpdateState cachedState = new UpdateState();

        public UpdateNotificationService(
            UpdateCheckOptions options,
            IUpdateManifestClient manifestClient,
            IUpdateStateStore stateStore,
            string currentVersion,
            Func<DateTime> getUtcNow = null)
        {
            this.options = options ?? throw new ArgumentNullException(nameof(options));
            this.manifestClient = manifestClient ?? throw new ArgumentNullException(nameof(manifestClient));
            this.stateStore = stateStore ?? throw new ArgumentNullException(nameof(stateStore));
            this.currentVersion = currentVersion ?? string.Empty;
            this.getUtcNow = getUtcNow ?? (() => DateTime.UtcNow);
            CurrentState = UpdateNotificationState.Empty;
        }

        public event EventHandler StateChanged;

        public UpdateNotificationState CurrentState { get; private set; }

        public void LoadCachedState()
        {
            ApplyState(LoadStateSafely());
            RaiseStateChanged(null);
        }

        public void StartBackgroundCheck(SynchronizationContext uiContext)
        {
            Task.Run(async () =>
            {
                try
                {
                    if (options.StartupDelay > TimeSpan.Zero)
                    {
                        await Task.Delay(options.StartupDelay).ConfigureAwait(false);
                    }

                    await CheckForUpdatesCoreAsync(CancellationToken.None, raiseStateChanged: false).ConfigureAwait(false);
                    RaiseStateChanged(uiContext);
                }
                catch (Exception ex)
                {
                    OfficeAgentLog.Error("updates", "background_check.failed", "Background update check failed.", ex);
                }
            });
        }

        public async Task CheckForUpdatesAsync(CancellationToken cancellationToken)
        {
            await CheckForUpdatesCoreAsync(cancellationToken, raiseStateChanged: true).ConfigureAwait(false);
        }

        private async Task CheckForUpdatesCoreAsync(CancellationToken cancellationToken, bool raiseStateChanged)
        {
            var previousState = LoadStateSafely();

            try
            {
                if (!options.IsEnabled || string.IsNullOrWhiteSpace(options.ManifestUrl))
                {
                    OfficeAgentLog.Info("updates", "check.skipped_disabled", "Update check skipped because update checks are disabled.");
                    ApplyState(previousState);
                    RaiseStateChangedIfNeeded(raiseStateChanged);
                    return;
                }

                var nowUtc = getUtcNow();
                if (IsCacheFresh(previousState, nowUtc))
                {
                    ApplyState(previousState);
                    RaiseStateChangedIfNeeded(raiseStateChanged);
                    return;
                }

                using (var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken))
                {
                    timeoutCts.CancelAfter(options.RequestTimeout);
                    var manifest = await manifestClient.GetManifestAsync(options.ManifestUrl, timeoutCts.Token).ConfigureAwait(false);
                    if (!UpdateVersionComparer.TryParse(manifest?.LatestVersion, out var latestVersion))
                    {
                        throw new InvalidOperationException("Update manifest latestVersion is invalid.");
                    }

                    var nextState = CloneState(previousState);
                    nextState.ApplyManifest(manifest, nowUtc);
                    MergeLatestIgnoredVersion(nextState);
                    stateStore.Save(nextState);
                    ApplyState(nextState);
                    RaiseStateChangedIfNeeded(raiseStateChanged);
                }
            }
            catch (Exception ex)
            {
                OfficeAgentLog.Warn("updates", "check.failed", "Update check failed.", ex.Message);
                ApplyState(previousState);
                RaiseStateChangedIfNeeded(raiseStateChanged);
            }
        }

        public void IgnoreCurrentVersion()
        {
            UpdateState nextState;
            lock (syncRoot)
            {
                if (string.IsNullOrWhiteSpace(cachedState.LatestVersion))
                {
                    return;
                }

                nextState = CloneState(cachedState);
                nextState.IgnoredVersion = nextState.LatestVersion;
                cachedState = nextState;
                CurrentState = BuildNotificationState(nextState);
            }

            try
            {
                stateStore.Save(nextState);
            }
            catch (Exception ex)
            {
                OfficeAgentLog.Warn("updates", "ignore.save_failed", "Failed to save ignored update version.", ex.Message);
            }

            OfficeAgentLog.Info("updates", "version.ignored", "User ignored the current update version.", nextState.IgnoredVersion);
            RaiseStateChanged(null);
        }

        private void MergeLatestIgnoredVersion(UpdateState state)
        {
            var latestState = LoadStateSafely();
            if (!string.IsNullOrWhiteSpace(latestState.IgnoredVersion))
            {
                state.IgnoredVersion = latestState.IgnoredVersion;
            }

            lock (syncRoot)
            {
                if (!string.IsNullOrWhiteSpace(cachedState.IgnoredVersion))
                {
                    state.IgnoredVersion = cachedState.IgnoredVersion;
                }
            }
        }

        private bool IsCacheFresh(UpdateState state, DateTime nowUtc)
        {
            return state.LastCheckedAtUtc.HasValue &&
                   state.LastCheckedAtUtc.Value <= nowUtc &&
                   nowUtc - state.LastCheckedAtUtc.Value < options.CacheDuration;
        }

        private UpdateState LoadStateSafely()
        {
            try
            {
                return stateStore.Load();
            }
            catch (Exception ex)
            {
                OfficeAgentLog.Warn("updates", "state.load_failed", "Failed to load update state.", ex.Message);
                return new UpdateState();
            }
        }

        private void ApplyState(UpdateState state)
        {
            lock (syncRoot)
            {
                cachedState = CloneState(state);
                CurrentState = BuildNotificationState(cachedState);
            }
        }

        private void RaiseStateChangedIfNeeded(bool raiseStateChanged)
        {
            if (!raiseStateChanged)
            {
                return;
            }

            RaiseStateChanged(null);
        }

        private UpdateNotificationState BuildNotificationState(UpdateState state)
        {
            if (state == null ||
                !UpdateVersionComparer.IsNewerThanCurrent(state.LatestVersion, currentVersion) ||
                string.Equals(state.LatestVersion, state.IgnoredVersion, StringComparison.OrdinalIgnoreCase))
            {
                return UpdateNotificationState.Empty;
            }

            return new UpdateNotificationState(
                true,
                state.LatestVersion,
                state.DownloadUrl,
                state.ReleaseNotesUrl,
                state.PublishedAtUtc,
                state.Title,
                state.Summary);
        }

        private void RaiseStateChanged(SynchronizationContext uiContext)
        {
            var handler = StateChanged;
            if (handler == null)
            {
                return;
            }

            if (uiContext == null)
            {
                InvokeStateChangedHandler(handler);
                return;
            }

            try
            {
                uiContext.Post(_ => InvokeStateChangedHandler(handler), null);
            }
            catch (Exception ex)
            {
                OfficeAgentLog.Warn("updates", "state_changed.post_failed", "Failed to post update notification state change.", ex.Message);
            }
        }

        private void InvokeStateChangedHandler(EventHandler handler)
        {
            foreach (EventHandler subscriber in handler.GetInvocationList())
            {
                try
                {
                    subscriber(this, EventArgs.Empty);
                }
                catch (Exception ex)
                {
                    OfficeAgentLog.Warn("updates", "state_changed.handler_failed", "Update notification state change subscriber failed.", ex.Message);
                }
            }
        }

        private static UpdateState CloneState(UpdateState state)
        {
            state = state ?? new UpdateState();
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
