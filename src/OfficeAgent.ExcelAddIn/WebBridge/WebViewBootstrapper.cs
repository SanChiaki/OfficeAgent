using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Newtonsoft.Json;
using OfficeAgent.Core.Diagnostics;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
using OfficeAgent.Infrastructure.Storage;

namespace OfficeAgent.ExcelAddIn.WebBridge
{
    internal sealed class WebViewBootstrapper
    {
        private const string VirtualHost = "appassets.officeagent.local";
        private readonly WebView2 webView;
        private readonly WebMessageRouter messageRouter;
        private bool isInitialized;
        private bool isProcessing;

        public WebViewBootstrapper(
            WebView2 webView,
            FileSessionStore sessionStore,
            FileSettingsStore settingsStore,
            IExcelContextService excelContextService,
            IExcelCommandExecutor excelCommandExecutor,
            IAgentOrchestrator agentOrchestrator)
        {
            this.webView = webView;
            messageRouter = new WebMessageRouter(sessionStore, settingsStore, excelContextService, excelCommandExecutor, agentOrchestrator);
        }

        public async Task InitializeAsync()
        {
            OfficeAgentLog.Info("webview", "initialize.begin", "Initializing WebView2.");
            var environment = await CoreWebView2Environment.CreateAsync(
                browserExecutableFolder: null,
                userDataFolder: GetUserDataFolder());

            await webView.EnsureCoreWebView2Async(environment);
            webView.CoreWebView2.WebMessageReceived += CoreWebView2_WebMessageReceived;
            isInitialized = true;

            var frontendFolder = ResolveFrontendFolder();
            if (frontendFolder == null)
            {
                OfficeAgentLog.Warn("webview", "frontend.missing", "Frontend assets were not found for the task pane.");
                webView.NavigateToString(BuildFallbackHtml());
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
                    PostErrorResponse(rawJson, "busy", "Another request is already in progress. Please wait.");
                    return;
                }

                isProcessing = true;
                try
                {
                    // await releases the UI thread during the HTTP call; ConfigureAwait(true) in the
                    // async chain keeps every continuation on the SynchronizationContext (UI thread)
                    // so COM calls in the orchestrator remain safe.
                    var responseJson = await messageRouter.RouteAsync(rawJson).ConfigureAwait(true);
                    webView.CoreWebView2.PostWebMessageAsJson(responseJson);
                }
                catch (Exception error)
                {
                    var requestId = ExtractRequestId(rawJson);
                    var message = error is OperationCanceledException
                        ? "Agent request timed out."
                        : (error.Message ?? "Agent execution failed.");
                    PostError(requestId, rawJson, "internal_error", message);
                }
                finally
                {
                    isProcessing = false;
                }

                return;
            }

            var syncResponse = messageRouter.Route(rawJson);
            webView.CoreWebView2.PostWebMessageAsJson(syncResponse);
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
            webView.CoreWebView2.PostWebMessageAsJson(messageJson);
        }

        private static bool IsLongRunningMessage(string rawJson)
        {
            try
            {
                var obj = JsonConvert.DeserializeObject<WebMessageRequest>(rawJson);
                return obj != null &&
                    (string.Equals(obj.Type, BridgeMessageTypes.RunAgent, StringComparison.Ordinal) ||
                     string.Equals(obj.Type, BridgeMessageTypes.RunSkill, StringComparison.Ordinal));
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

        private void PostErrorResponse(string rawJson, string code, string message)
        {
            var requestId = ExtractRequestId(rawJson);
            var requestType = "bridge.unknown";
            try
            {
                var obj = JsonConvert.DeserializeObject<WebMessageRequest>(rawJson);
                if (obj != null) requestType = obj.Type ?? requestType;
            }
            catch { }

            PostError(requestId, rawJson, code, message);
        }

        private void PostError(string requestId, string rawJson, string code, string message)
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
            var errorJson = JsonConvert.SerializeObject(errorResponse, new Newtonsoft.Json.Serialization.CamelCasePropertyNamesContractResolver() != null
                ? new JsonSerializerSettings
                {
                    ContractResolver = new Newtonsoft.Json.Serialization.CamelCasePropertyNamesContractResolver(),
                    NullValueHandling = NullValueHandling.Ignore,
                }
                : null);
            webView.CoreWebView2.PostWebMessageAsJson(errorJson);
        }

        private static string BuildFallbackHtml()
        {
            return @"<!doctype html>
<html lang=""en"">
  <head>
    <meta charset=""utf-8"" />
    <title>OfficeAgent</title>
    <style>
      body { font-family: Segoe UI, sans-serif; padding: 24px; color: #1f2937; }
      code { background: #f3f4f6; padding: 2px 6px; border-radius: 4px; }
    </style>
  </head>
  <body>
    <h1>OfficeAgent</h1>
    <p>Frontend assets were not found.</p>
    <p>Build <code>src/OfficeAgent.Frontend</code> and reopen the task pane.</p>
  </body>
</html>";
        }
    }
}
