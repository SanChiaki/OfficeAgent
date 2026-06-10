using System;
using System.Diagnostics;
using System.Text;
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

    internal static class AboutDialog
    {
        private const int DialogWidth = 540;
        private const int MinimumDialogHeight = 220;
        private const int MaximumDialogHeight = 380;
        private const int PromptVerticalPadding = 110;
        private const int PromptLineHeight = 20;

        public static AboutDialogAction ShowDialogForUpdate(AboutDialogModel model, HostLocalizedStrings strings)
        {
            model = model ?? new AboutDialogModel();
            strings = strings ?? HostLocalizedStrings.ForLocale("en");
            var owner = ExcelDialogOwner.FromCurrentApplication();
            var message = CreateMessage(model, strings);
            var action = AboutDialogAction.Close;
            var buttons = CreateButtons(model, strings, result => action = result);
            TemplatePromptDialog.ShowPrompt(
                owner,
                strings.RibbonAboutDialogTitle,
                message,
                MessageBoxIcon.Information,
                new TemplatePromptDialog.PromptOptions
                {
                    Width = DialogWidth,
                    Height = EstimatePromptHeight(message),
                    EnableMessageScroll = true,
                },
                buttons);
            return action;
        }

        private static string CreateMessage(AboutDialogModel model, HostLocalizedStrings strings)
        {
            var builder = new StringBuilder()
                .AppendLine($"{strings.AboutCurrentVersionLabel}: {model.AppVersion}")
                .AppendLine($"{strings.AboutAssemblyVersionLabel}: {model.AssemblyVersion}")
                .AppendLine($"{strings.AboutBuildConfigurationLabel}: {model.BuildConfiguration}")
                .AppendLine($"{strings.AboutBuildTimeLabel}: {model.BuildTime}")
                .AppendLine();

            if (model.HasNewVersion)
            {
                builder
                    .AppendLine(strings.AboutNewVersionAvailableTitle)
                    .AppendLine($"{strings.AboutLatestVersionLabel}: {model.LatestVersion}");
                if (model.PublishedAtUtc.HasValue)
                {
                    builder.AppendLine($"{strings.AboutPublishedAtLabel}: {model.PublishedAtUtc.Value.ToLocalTime():yyyy-MM-dd HH:mm:ss}");
                }

                if (!string.IsNullOrWhiteSpace(model.UpdateTitle))
                {
                    builder.AppendLine(model.UpdateTitle.Trim());
                }

                if (!string.IsNullOrWhiteSpace(model.UpdateSummary))
                {
                    builder.AppendLine(model.UpdateSummary.Trim());
                }
            }
            else
            {
                builder.AppendLine(strings.AboutNoUpdateAvailableText);
            }

            return builder.ToString().TrimEnd();
        }

        private static TemplatePromptDialog.DialogButtonSpec[] CreateButtons(
            AboutDialogModel model,
            HostLocalizedStrings strings,
            Action<AboutDialogAction> setAction)
        {
            if (!model.HasNewVersion)
            {
                return new[]
                {
                    new TemplatePromptDialog.DialogButtonSpec(strings.CloseButtonText, DialogResult.Cancel, isCancel: true),
                };
            }

            if (!IsSupportedHttpUrl(model.DownloadUrl))
            {
                return new[]
                {
                    new TemplatePromptDialog.DialogButtonSpec(strings.CloseButtonText, DialogResult.Cancel, isCancel: true),
                    new TemplatePromptDialog.DialogButtonSpec(strings.AboutIgnoreVersionButtonText, DialogResult.OK, action: owner => setAction(AboutDialogAction.IgnoreVersion)),
                };
            }

            return new[]
            {
                new TemplatePromptDialog.DialogButtonSpec(strings.CloseButtonText, DialogResult.Cancel, isCancel: true),
                new TemplatePromptDialog.DialogButtonSpec(strings.AboutIgnoreVersionButtonText, DialogResult.OK, action: owner => setAction(AboutDialogAction.IgnoreVersion)),
                new TemplatePromptDialog.DialogButtonSpec(strings.AboutDownloadButtonText, DialogResult.None, action: owner => OpenUrl(owner, model.DownloadUrl, strings)),
            };
        }

        private static int EstimatePromptHeight(string message)
        {
            var lineCount = 1;
            foreach (var character in message ?? string.Empty)
            {
                if (character == '\n')
                {
                    lineCount++;
                }
            }

            var desiredHeight = PromptVerticalPadding + (lineCount * PromptLineHeight);
            return Math.Max(MinimumDialogHeight, Math.Min(MaximumDialogHeight, desiredHeight));
        }

        private static bool IsSupportedHttpUrl(string url)
        {
            return Uri.TryCreate(url, UriKind.Absolute, out var uri) &&
                   (string.Equals(uri.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase));
        }

        private static void OpenUrl(IWin32Window owner, string url, HostLocalizedStrings strings)
        {
            if (!IsSupportedHttpUrl(url))
            {
                MessageBox.Show(owner, strings.AboutOpenUrlFailedMessage(url), strings.HostWindowTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                MessageBox.Show(owner, strings.AboutOpenUrlFailedMessage(ex.Message), strings.HostWindowTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
