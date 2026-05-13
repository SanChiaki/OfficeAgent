using System;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using OfficeAgent.Core.Analytics;
using OfficeAgent.Core.Diagnostics;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
using OfficeAgent.ExcelAddIn.Localization;
using OfficeAgent.ExcelAddIn.WebBridge;
using OfficeAgent.Infrastructure.Http;
using OfficeAgent.Infrastructure.Storage;

namespace OfficeAgent.ExcelAddIn.TaskPane
{
    internal sealed class TaskPaneHostControl : UserControl
    {
        private readonly WebView2 webView;
        private readonly WebViewBootstrapper bootstrapper;
        private SelectionContext pendingSelectionContext;
        private bool isLoadStarted;
        private bool isBridgeReady;
        private readonly Func<AppSettings, string> getResolvedUiLocale;
        private readonly FileSettingsStore settingsStore;

        public TaskPaneHostControl(
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
            this.settingsStore = settingsStore;
            this.getResolvedUiLocale = getResolvedUiLocale ?? throw new ArgumentNullException(nameof(getResolvedUiLocale));
            Dock = DockStyle.Fill;

            webView = new WebView2
            {
                Dock = DockStyle.Fill
            };
            Controls.Add(webView);

            bootstrapper = new WebViewBootstrapper(webView, sessionStore, settingsStore, excelContextService, excelCommandExecutor, agentOrchestrator, sharedCookies, cookieStore, getResolvedUiLocale, analyticsService);
            Load += TaskPaneHostControl_Load;
        }

        private async void TaskPaneHostControl_Load(object sender, EventArgs e)
        {
            if (isLoadStarted || DesignMode)
            {
                return;
            }

            isLoadStarted = true;

            try
            {
                await bootstrapper.InitializeAsync();
                isBridgeReady = true;
                OfficeAgentLog.Info("webview", "initialized", "WebView2 task pane initialized.");
                if (pendingSelectionContext != null)
                {
                    bootstrapper.PublishSelectionContext(pendingSelectionContext);
                    pendingSelectionContext = null;
                }
            }
            catch (WebView2RuntimeNotFoundException)
            {
                webView.Visible = false;
                OfficeAgentLog.Warn("webview", "runtime.missing", "WebView2 Runtime is missing.");
                Controls.Add(new Label
                {
                    Dock = DockStyle.Fill,
                    Text = GetStrings().TaskPaneRuntimeMissingMessage,
                    TextAlign = System.Drawing.ContentAlignment.MiddleCenter
                });
            }
            catch (Exception error)
            {
                webView.Visible = false;
                OfficeAgentLog.Error("webview", "initialize.failed", "WebView2 initialization failed.", error);
                Controls.Add(new Label
                {
                    Dock = DockStyle.Fill,
                    Text = GetStrings().TaskPaneInitializationFailedMessage,
                    TextAlign = System.Drawing.ContentAlignment.MiddleCenter
                });
            }
        }

        public void PublishSelectionContext(SelectionContext selectionContext)
        {
            if (!isBridgeReady)
            {
                pendingSelectionContext = selectionContext;
                return;
            }

            bootstrapper.PublishSelectionContext(selectionContext);
        }

        private HostLocalizedStrings GetStrings()
        {
            return HostLocalizedStrings.ForLocale(getResolvedUiLocale(settingsStore.Load()));
        }
    }
}
