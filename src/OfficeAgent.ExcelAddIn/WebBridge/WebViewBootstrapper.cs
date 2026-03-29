using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Newtonsoft.Json;
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

        public WebViewBootstrapper(
            WebView2 webView,
            FileSessionStore sessionStore,
            FileSettingsStore settingsStore,
            IExcelContextService excelContextService)
        {
            this.webView = webView;
            messageRouter = new WebMessageRouter(sessionStore, settingsStore, excelContextService);
        }

        public async Task InitializeAsync()
        {
            var environment = await CoreWebView2Environment.CreateAsync(
                browserExecutableFolder: null,
                userDataFolder: GetUserDataFolder());

            await webView.EnsureCoreWebView2Async(environment);
            webView.CoreWebView2.WebMessageReceived += CoreWebView2_WebMessageReceived;
            isInitialized = true;

            var frontendFolder = ResolveFrontendFolder();
            if (frontendFolder == null)
            {
                webView.NavigateToString(BuildFallbackHtml());
                return;
            }

            webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                VirtualHost,
                frontendFolder,
                CoreWebView2HostResourceAccessKind.Allow);

            webView.Source = new Uri($"https://{VirtualHost}/index.html");
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

        private void CoreWebView2_WebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            var responseJson = messageRouter.Route(e.WebMessageAsJson);
            webView.CoreWebView2.PostWebMessageAsJson(responseJson);
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
