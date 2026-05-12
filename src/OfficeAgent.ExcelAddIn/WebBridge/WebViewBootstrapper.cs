using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Newtonsoft.Json;
using OfficeAgent.Core.Analytics;
using OfficeAgent.Core.Diagnostics;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
using OfficeAgent.ExcelAddIn.Localization;
using OfficeAgent.Infrastructure.Http;
using OfficeAgent.Infrastructure.Storage;

namespace OfficeAgent.ExcelAddIn.WebBridge
{
    internal sealed class WebViewBootstrapper
    {
        private const string VirtualHost = "appassets.officeagent.local";
        private readonly WebView2 webView;
        private readonly WebMessageRouter messageRouter;
        private readonly FileSettingsStore settingsStore;
        private readonly Func<AppSettings, string> getResolvedUiLocale;
        private bool isInitialized;
        private bool isProcessing;

        public WebViewBootstrapper(
            WebView2 webView,
            FileSessionStore sessionStore,
            FileSettingsStore settingsStore,
            IExcelContextService excelContextService,
            IExcelCommandExecutor excelCommandExecutor,
            IAgentOrchestrator agentOrchestrator,
            SharedCookieContainer sharedCookies,
            FileCookieStore cookieStore,
            Func<AppSettings, string> getResolvedUiLocale,
            IAnalyticsService analyticsService = null)
        {
            this.webView = webView;
            this.settingsStore = settingsStore;
            this.getResolvedUiLocale = getResolvedUiLocale ?? throw new ArgumentNullException(nameof(getResolvedUiLocale));
            messageRouter = new WebMessageRouter(sessionStore, settingsStore, excelContextService, excelCommandExecutor, agentOrchestrator, sharedCookies, cookieStore, getResolvedUiLocale, analyticsService);
        }

        public async Task InitializeAsync()
        {
            // VSTO does not call Application.Run, so WindowsFormsSynchronizationContext is
            // not installed by default on the Excel STA thread. Without it, ConfigureAwait(true)
            // in the async chain has no SynchronizationContext to capture and continuations land
            // on thread-pool threads instead of the UI thread, breaking COM calls and WebView2
            // access. Install it here, before the first await, while we are still on the UI thread.
            if (!(SynchronizationContext.Current is WindowsFormsSynchronizationContext))
            {
                SynchronizationContext.SetSynchronizationContext(
                    new WindowsFormsSynchronizationContext());
            }

            OfficeAgentLog.Info("webview", "initialize.begin", "Initializing WebView2.");
            var environment = await CoreWebView2Environment.CreateAsync(
                browserExecutableFolder: null,
                userDataFolder: GetUserDataFolder());

            await webView.EnsureCoreWebView2Async(environment);
            webView.CoreWebView2.WebMessageReceived += CoreWebView2_WebMessageReceived;
            webView.CoreWebView2.ProcessFailed += CoreWebView2_ProcessFailed;
            isInitialized = true;

            var frontendFolder = ResolveFrontendFolder();
            if (frontendFolder == null)
            {
                OfficeAgentLog.Warn("webview", "frontend.missing", "Frontend assets were not found for the task pane.");
                webView.NavigateToString(GetStrings().BootstrapperFallbackHtml);
                return;
            }

            webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                VirtualHost,
                frontendFolder,
                CoreWebView2HostResourceAccessKind.Allow);

            webView.Source = new Uri($"https://{VirtualHost}/index.html");
            OfficeAgentLog.Info("webview", "navigate.index", "Navigated WebView2 to the packaged frontend.");
        }

        private static string GetUserDataFolder()
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OfficeAgent",
                "WebView2");
        }

        private static string ResolveFrontendFolder()
        {
            var installedFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "frontend");
            if (File.Exists(Path.Combine(installedFolder, "index.html")))
            {
                return installedFolder;
            }

            var developmentFolder = Path.GetFullPath(
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\OfficeAgent.Frontend\dist"));
            if (File.Exists(Path.Combine(developmentFolder, "index.html")))
            {
                return developmentFolder;
            }

            return null;
        }

        private async void CoreWebView2_WebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            var rawJson = e.WebMessageAsJson;

            if (IsLongRunningMessage(rawJson))
            {
                if (isProcessing)
                {
                    // Swallow any posting failure — the process might be in a bad state.
                    TryPostErrorResponse(rawJson, "busy", GetStrings().BridgeBusyMessage);
                    return;
                }

                isProcessing = true;
                try
                {
                    // await releases the UI thread during the HTTP call; ConfigureAwait(true) in the
                    // async chain keeps every continuation on the SynchronizationContext (UI thread)
                    // so COM calls in the orchestrator remain safe.
                    var responseJson = await messageRouter.RouteAsync(rawJson).ConfigureAwait(true);
                    TryPostWebMessage(responseJson);
                }
                catch (Exception error)
                {
                    // An exception here means RouteAsync itself faulted (unusual path).
                    // Use TryPostError so a secondary failure cannot escape async void.
                    var requestId = ExtractRequestId(rawJson);
                    var message = error is OperationCanceledException
                        ? GetStrings().BridgeAgentRequestTimedOutMessage
                        : (error.Message ?? GetStrings().BridgeAgentExecutionFailedMessage);
                    TryPostError(requestId, rawJson, "internal_error", message);
                }
                finally
                {
                    isProcessing = false;
                }

                return;
            }

            // Synchronous fast-path for non-LLM messages.
            try
            {
                var syncResponse = messageRouter.Route(rawJson);
                TryPostWebMessage(syncResponse);
            }
            catch (Exception error)
            {
                OfficeAgentLog.Error("bridge", "sync.failed", "Sync bridge message failed.", error);
            }
        }

        public void PublishSelectionContext(SelectionContext selectionContext)
        {
            if (!isInitialized || webView.CoreWebView2 == null || selectionContext == null)
            {
                return;
            }

            var messageJson = JsonConvert.SerializeObject(new WebMessageEvent
            {
                Type = BridgeMessageTypes.SelectionContextChanged,
                Payload = selectionContext,
            });

            // Called from Application.SheetSelectionChange (a COM event handler).
            // If the WebView2 renderer has crashed or is in a bad state, PostWebMessageAsJson
            // can throw. Swallow the exception here to prevent it from propagating back to
            // Excel's COM event dispatcher, which would crash the host process.
            try
            {
                webView.CoreWebView2.PostWebMessageAsJson(messageJson);
            }
            catch (Exception error)
            {
                OfficeAgentLog.Warn("webview", "publish.failed", $"Failed to publish selection context: {error.Message}");
            }
        }

        private static void CoreWebView2_ProcessFailed(object sender, CoreWebView2ProcessFailedEventArgs e)
        {
            OfficeAgentLog.Warn("webview", "process.failed", $"WebView2 process failed: {e.ProcessFailedKind}.");
        }

        private static bool IsLongRunningMessage(string rawJson)
        {
            try
            {
                var obj = JsonConvert.DeserializeObject<WebMessageRequest>(rawJson);
                return obj != null &&
                    (string.Equals(obj.Type, BridgeMessageTypes.RunAgent, StringComparison.Ordinal) ||
                     string.Equals(obj.Type, BridgeMessageTypes.RunSkill, StringComparison.Ordinal) ||
                     string.Equals(obj.Type, BridgeMessageTypes.Login, StringComparison.Ordinal));
            }
            catch
            {
                return false;
            }
        }

        private static string ExtractRequestId(string rawJson)
        {
            try
            {
                var obj = JsonConvert.DeserializeObject<WebMessageRequest>(rawJson);
                return obj?.RequestId ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private void TryPostWebMessage(string json)
        {
            try
            {
                // Belt-and-suspenders: if the continuation somehow ran on a thread-pool thread
                // despite ConfigureAwait(true), marshal explicitly to the control's UI thread.
                if (webView.InvokeRequired)
                {
                    webView.Invoke((Action)(() => webView.CoreWebView2?.PostWebMessageAsJson(json)));
                }
                else
                {
                    webView.CoreWebView2?.PostWebMessageAsJson(json);
                }
            }
            catch (Exception error)
            {
                OfficeAgentLog.Warn("webview", "post.failed", $"PostWebMessage failed: {error.Message}");
            }
        }

        private void TryPostErrorResponse(string rawJson, string code, string message)
        {
            var requestId = ExtractRequestId(rawJson);
            TryPostError(requestId, rawJson, code, message);
        }

        private void TryPostError(string requestId, string rawJson, string code, string message)
        {
            var requestType = "bridge.unknown";
            try
            {
                var obj = JsonConvert.DeserializeObject<WebMessageRequest>(rawJson);
                if (obj != null) requestType = obj.Type ?? requestType;
            }
            catch { }

            var errorResponse = new WebMessageResponse
            {
                Type = requestType,
                RequestId = requestId,
                Ok = false,
                Error = new WebMessageError
                {
                    Code = code,
                    Message = message,
                },
            };

            var errorJson = JsonConvert.SerializeObject(errorResponse, new JsonSerializerSettings
            {
                ContractResolver = new Newtonsoft.Json.Serialization.CamelCasePropertyNamesContractResolver(),
                NullValueHandling = NullValueHandling.Ignore,
            });

            TryPostWebMessage(errorJson);
        }

        private HostLocalizedStrings GetStrings()
        {
            return HostLocalizedStrings.ForLocale(getResolvedUiLocale(settingsStore.Load() ?? new AppSettings()));
        }
    }
}
