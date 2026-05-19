using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using OfficeAgent.ExcelAddIn.Localization;

namespace OfficeAgent.ExcelAddIn.Dialogs
{
    internal enum AboutDialogAction
    {
        Close,
        IgnoreVersion,
    }

    internal sealed class AboutDialogModel
    {
        public string AppVersion { get; set; } = string.Empty;

        public string AssemblyVersion { get; set; } = string.Empty;

        public string BuildConfiguration { get; set; } = string.Empty;

        public string BuildTime { get; set; } = string.Empty;

        public bool HasNewVersion { get; set; }

        public string LatestVersion { get; set; } = string.Empty;

        public string DownloadUrl { get; set; } = string.Empty;

        public string ReleaseNotesUrl { get; set; } = string.Empty;

        public DateTime? PublishedAtUtc { get; set; }

        public string UpdateTitle { get; set; } = string.Empty;

        public string UpdateSummary { get; set; } = string.Empty;
    }

    internal sealed class AboutDialog : Form
    {
        private const int DialogWidth = 460;
        private const int HorizontalPadding = 18;
        private const int ButtonHeight = 28;
        private const int WrappedLineHeight = 96;

        private readonly AboutDialogModel model;
        private readonly HostLocalizedStrings strings;
        private AboutDialogAction action = AboutDialogAction.Close;

        private AboutDialog(AboutDialogModel model, HostLocalizedStrings strings)
        {
            this.model = model ?? new AboutDialogModel();
            this.strings = strings ?? HostLocalizedStrings.ForLocale("en");
            BuildLayout();
        }

        public static AboutDialogAction ShowDialogForUpdate(AboutDialogModel model, HostLocalizedStrings strings)
        {
            var owner = ExcelDialogOwner.FromCurrentApplication();
            using (var dialog = new AboutDialog(model, strings))
            {
                if (owner == null)
                {
                    dialog.ShowDialog();
                }
                else
                {
                    dialog.ShowDialog(owner);
                }

                return dialog.action;
            }
        }

        private void BuildLayout()
        {
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Dpi;
            Text = strings.RibbonAboutDialogTitle;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;

            var top = 16;
            AddLine($"{strings.AboutCurrentVersionLabel}: {model.AppVersion}", FontStyle.Regular, ref top);
            AddLine($"{strings.AboutAssemblyVersionLabel}: {model.AssemblyVersion}", FontStyle.Regular, ref top);
            AddLine($"{strings.AboutBuildConfigurationLabel}: {model.BuildConfiguration}", FontStyle.Regular, ref top);
            AddLine($"{strings.AboutBuildTimeLabel}: {model.BuildTime}", FontStyle.Regular, ref top);

            top += 8;
            if (model.HasNewVersion)
            {
                AddLine(strings.AboutNewVersionAvailableTitle, FontStyle.Bold, ref top);
                AddLine($"{strings.AboutLatestVersionLabel}: {model.LatestVersion}", FontStyle.Regular, ref top);
                if (model.PublishedAtUtc.HasValue)
                {
                    AddLine($"{strings.AboutPublishedAtLabel}: {model.PublishedAtUtc.Value.ToLocalTime():yyyy-MM-dd HH:mm:ss}", FontStyle.Regular, ref top);
                }

                if (!string.IsNullOrWhiteSpace(model.UpdateTitle))
                {
                    AddWrappedLine(model.UpdateTitle, ref top);
                }

                if (!string.IsNullOrWhiteSpace(model.UpdateSummary))
                {
                    AddWrappedLine(model.UpdateSummary, ref top);
                }
            }
            else
            {
                AddLine(strings.AboutNoUpdateAvailableText, FontStyle.Regular, ref top);
            }

            top += 10;
            AddButtons(top);
            ClientSize = new Size(DialogWidth, top + ButtonHeight + 18);
        }

        private void AddLine(string text, FontStyle style, ref int top)
        {
            var label = new Label
            {
                AutoSize = false,
                Text = text ?? string.Empty,
                Font = new Font(Font, style),
                Bounds = new Rectangle(HorizontalPadding, top, DialogWidth - (HorizontalPadding * 2), 22),
            };
            Controls.Add(label);
            top += 22;
        }

        private void AddWrappedLine(string text, ref int top)
        {
            var label = new Label
            {
                AutoSize = false,
                Text = text ?? string.Empty,
                Bounds = new Rectangle(HorizontalPadding, top, DialogWidth - (HorizontalPadding * 2), WrappedLineHeight),
            };
            Controls.Add(label);
            top += WrappedLineHeight + 2;
        }

        private void AddButtons(int top)
        {
            var right = DialogWidth - HorizontalPadding;
            var closeButton = CreateButton(strings.CloseButtonText, right - 76, top, 76);
            closeButton.DialogResult = DialogResult.Cancel;
            closeButton.Click += (sender, e) =>
            {
                action = AboutDialogAction.Close;
                Close();
            };
            Controls.Add(closeButton);
            right -= 84;

            if (model.HasNewVersion)
            {
                var ignoreButton = CreateButton(strings.AboutIgnoreVersionButtonText, right - 118, top, 118);
                ignoreButton.Click += (sender, e) =>
                {
                    action = AboutDialogAction.IgnoreVersion;
                    Close();
                };
                Controls.Add(ignoreButton);
                right -= 126;
            }

            if (IsSupportedHttpUrl(model.DownloadUrl))
            {
                var downloadButton = CreateButton(strings.AboutDownloadButtonText, right - 88, top, 88);
                downloadButton.Click += (sender, e) => OpenUrl(model.DownloadUrl);
                Controls.Add(downloadButton);
            }

            CancelButton = closeButton;
        }

        private static Button CreateButton(string text, int left, int top, int width)
        {
            return new Button
            {
                Text = text ?? string.Empty,
                Bounds = new Rectangle(Math.Max(HorizontalPadding, left), top, width, ButtonHeight),
            };
        }

        private static bool IsSupportedHttpUrl(string url)
        {
            return Uri.TryCreate(url, UriKind.Absolute, out var uri) &&
                   (string.Equals(uri.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase));
        }

        private void OpenUrl(string url)
        {
            if (!IsSupportedHttpUrl(url))
            {
                MessageBox.Show(this, strings.AboutOpenUrlFailedMessage(url), strings.HostWindowTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = url,
                    UseShellExecute = true,
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, strings.AboutOpenUrlFailedMessage(ex.Message), strings.HostWindowTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
