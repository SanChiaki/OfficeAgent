using System;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
using OfficeAgent.ExcelAddIn.WebBridge;
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

        public TaskPaneHostControl(
            FileSessionStore sessionStore,
            FileSettingsStore settingsStore,
            IExcelContextService excelContextService)
        {
            Dock = DockStyle.Fill;

            webView = new WebView2
            {
                Dock = DockStyle.Fill
            };
            Controls.Add(webView);

            bootstrapper = new WebViewBootstrapper(webView, sessionStore, settingsStore, excelContextService);
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
                if (pendingSelectionContext != null)
                {
                    bootstrapper.PublishSelectionContext(pendingSelectionContext);
                    pendingSelectionContext = null;
                }
            }
            catch (WebView2RuntimeNotFoundException)
            {
                webView.Visible = false;
                Controls.Add(new Label
                {
                    Dock = DockStyle.Fill,
                    Text = "WebView2 Runtime is required to render OfficeAgent.",
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
    }
}
