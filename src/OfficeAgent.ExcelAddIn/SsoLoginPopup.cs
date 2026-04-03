using System;
using System.IO;
using System.Net;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using OfficeAgent.Core.Diagnostics;
using OfficeAgent.Infrastructure.Http;
using OfficeAgent.Infrastructure.Storage;

namespace OfficeAgent.ExcelAddIn
{
    internal sealed class SsoLoginPopup : Form
    {
        private readonly string ssoUrl;
        private readonly string loginSuccessPath;
        private readonly SharedCookieContainer sharedCookies;
        private readonly FileCookieStore cookieStore;
        private WebView2 webView;

        public SsoLoginPopup(string ssoUrl, string loginSuccessPath, SharedCookieContainer sharedCookies, FileCookieStore cookieStore)
        {
            this.ssoUrl = ssoUrl ?? throw new ArgumentNullException(nameof(ssoUrl));
            this.loginSuccessPath = loginSuccessPath ?? string.Empty;
            this.sharedCookies = sharedCookies ?? throw new ArgumentNullException(nameof(sharedCookies));
            this.cookieStore = cookieStore ?? throw new ArgumentNullException(nameof(cookieStore));

            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = true;
            MinimizeBox = true;
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Resy AI - \u767B\u5F55";
            Size = new System.Drawing.Size(1024, 700);
            MinimumSize = new System.Drawing.Size(600, 400);

            var cancelButton = new Button
            {
                Text = "\u53D6\u6D88",
                Dock = DockStyle.Bottom,
                Height = 36,
            };
            cancelButton.Click += (sender, e) =>
            {
                DialogResult = DialogResult.Cancel;
                Close();
            };

            webView = new WebView2
            {
                Dock = DockStyle.Fill,
            };

            // Add cancel button first so it takes bottom space, then WebView2 fills the rest.
            Controls.Add(cancelButton);
            Controls.Add(webView);
        }

        /// <summary>
        /// Initializes the WebView2 control and navigates to the SSO URL.
        /// Must be called on the UI thread before ShowDialog().
        /// </summary>
        public async Task InitializeAsync()
        {
            var userDataFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OfficeAgent",
                "webview2-sso");

            var environment = await CoreWebView2Environment.CreateAsync(
                browserExecutableFolder: null,
                userDataFolder: userDataFolder);

            await webView.EnsureCoreWebView2Async(environment);
            webView.CoreWebView2.NavigationCompleted += CoreWebView2_NavigationCompleted;
            webView.CoreWebView2.Navigate(ssoUrl);

            OfficeAgentLog.Info("sso", "popup.navigating", "SSO login popup navigating.", ssoUrl);
        }

        private void CoreWebView2_NavigationCompleted(object sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            if (!e.IsSuccess)
            {
                return;
            }

            try
            {
                // If a login success path is configured (e.g. /rest/login), check if the
                // current URL's path contains it AND the page loaded successfully.
                if (!string.IsNullOrWhiteSpace(loginSuccessPath))
                {
                    var currentUri = new Uri(webView.CoreWebView2.Source);
                    var currentPath = currentUri.AbsolutePath.TrimEnd('/');
                    var normalizedPath = loginSuccessPath.Trim().TrimEnd('/');

                    if (!string.IsNullOrWhiteSpace(normalizedPath) &&
                        currentPath.IndexOf(normalizedPath, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        OfficeAgentLog.Info(
                            "sso", "login.success_marker",
                            $"SSO login detected via success path '{loginSuccessPath}'.", currentUri.AbsoluteUri);

                        CaptureCookies();
                        DialogResult = DialogResult.OK;
                        Close();
                        return;
                    }
                }
            }
            catch (Exception error)
            {
                OfficeAgentLog.Error("sso", "cookie.capture.failed", "Failed to capture SSO cookies.", error);
            }
        }

        private async void CaptureCookies()
        {
            try
            {
                var ssoAuthority = new Uri(ssoUrl).Authority;
                var cookies = await webView.CoreWebView2.CookieManager.GetCookiesAsync(ssoUrl);

                foreach (var cookie in cookies)
                {
                    var netCookie = new System.Net.Cookie(cookie.Name, cookie.Value, cookie.Path, cookie.Domain)
                    {
                        Secure = cookie.IsSecure,
                        HttpOnly = cookie.IsHttpOnly,
                    };

                    if (cookie.Expires != DateTime.MinValue)
                    {
                        netCookie.Expires = cookie.Expires;
                    }

                    sharedCookies.Container.Add(netCookie);
                }

                cookieStore.Save(sharedCookies.Container, ssoAuthority);

                OfficeAgentLog.Info("sso", "login.succeeded", "SSO login completed, cookies captured.", ssoAuthority);
            }
            catch (Exception error)
            {
                OfficeAgentLog.Error("sso", "cookie.capture.failed", "Failed to capture SSO cookies.", error);
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                webView?.Dispose();
            }

            base.Dispose(disposing);
        }
    }
}
