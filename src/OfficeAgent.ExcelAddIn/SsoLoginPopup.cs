using System;
using System.Drawing;
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
        private bool hasLoginSucceeded;

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

            var btnPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 44,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(12, 4, 12, 4),
            };

            var loginOkButton = new Button
            {
                Text = "\u5DF2\u767B\u5F55",
                Width = 90,
                Height = 34,
            };
            loginOkButton.Click += (sender, e) =>
            {
                OfficeAgentLog.Info("sso", "login.manual_confirm", "User clicked \"已登录\" button; capturing cookies.");
                CaptureCookiesAndClose();
            };

            var cancelButton = new Button
            {
                Text = "\u53D6\u6D88",
                Width = 90,
                Height = 34,
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

            btnPanel.Controls.Add(cancelButton);
            btnPanel.Controls.Add(loginOkButton);
            Controls.Add(webView);
            Controls.Add(btnPanel);
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

            if (!string.IsNullOrWhiteSpace(loginSuccessPath))
            {
                webView.CoreWebView2.WebResourceResponseReceived += CoreWebView2_WebResourceResponseReceived;
                webView.CoreWebView2.NavigationCompleted += CoreWebView2_NavigationCompleted;
            }

            webView.CoreWebView2.Navigate(ssoUrl);

            OfficeAgentLog.Info("sso", "popup.navigating", "SSO login popup navigating.", ssoUrl);
        }

        private void CoreWebView2_WebResourceResponseReceived(object sender, CoreWebView2WebResourceResponseReceivedEventArgs e)
        {
            if (hasLoginSucceeded)
            {
                return;
            }

            try
            {
                var uri = e.Request?.Uri;
                if (string.IsNullOrEmpty(uri))
                {
                    return;
                }

                var requestUri = new Uri(uri);
                var marker = loginSuccessPath.Trim();
                var statusCode = e.Response?.StatusCode;

                OfficeAgentLog.Info(
                    "sso", "login.response_seen",
                    $"SSO response: {requestUri.AbsolutePath} => {(int?)statusCode ?? 0}  marker='{marker}'");

                if (!string.IsNullOrEmpty(marker))
                {
                    var idx = requestUri.AbsolutePath.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
                    if (idx < 0)
                    {
                        OfficeAgentLog.Info(
                            "sso", "login.path_no_match",
                            $"Path '{requestUri.AbsolutePath}' does not contain marker '{marker}'.");
                        return;
                    }

                    OfficeAgentLog.Info(
                        "sso", "login.path_matched",
                        $"Path '{requestUri.AbsolutePath}' contains marker '{marker}' at index {idx}.");
                }

                if (statusCode == null || statusCode != 200)
                {
                    OfficeAgentLog.Info(
                        "sso", "login.status_no_match",
                        $"Status {(int?)statusCode ?? -1} is not 200; not marking login successful.");
                    return;
                }

                MarkLoginSucceeded();
            }
            catch (Exception error)
            {
                OfficeAgentLog.Error("sso", "response.detection_failed", "Error during login response detection.", error);
            }
        }

        private void CoreWebView2_NavigationCompleted(object sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            if (hasLoginSucceeded || string.IsNullOrWhiteSpace(loginSuccessPath))
            {
                return;
            }

            try
            {
                var currentUri = webView.CoreWebView2.Source;
                if (string.IsNullOrEmpty(currentUri))
                {
                    return;
                }

                var marker = loginSuccessPath.Trim();
                var uri = new Uri(currentUri);

                OfficeAgentLog.Info(
                    "sso", "login.nav_completed",
                    $"Navigation completed: {currentUri}, marker='{marker}', success={e.IsSuccess}");

                if (!string.IsNullOrEmpty(marker) &&
                    uri.AbsolutePath.IndexOf(marker, StringComparison.OrdinalIgnoreCase) >= 0 &&
                    e.IsSuccess)
                {
                    OfficeAgentLog.Info(
                        "sso", "login.nav_path_matched",
                        $"Navigation path '{uri.AbsolutePath}' contains marker '{marker}'; marking login successful.");
                    MarkLoginSucceeded();
                }
            }
            catch (Exception error)
            {
                OfficeAgentLog.Error("sso", "nav.detection_failed", "Error during navigation detection.", error);
            }
        }

        private void MarkLoginSucceeded()
        {
            if (hasLoginSucceeded)
            {
                return;
            }

            hasLoginSucceeded = true;
            OfficeAgentLog.Info("sso", "login.response_detected", "Login success detected; capturing cookies and closing.");
            BeginInvoke(new Action(CaptureCookiesAndClose));
        }

        private async void CaptureCookiesAndClose()
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

            DialogResult = DialogResult.OK;
            Close();
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
