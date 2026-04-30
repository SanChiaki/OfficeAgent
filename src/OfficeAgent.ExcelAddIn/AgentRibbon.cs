using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using OfficeAgent.Core;
using OfficeAgent.Core.Diagnostics;
using OfficeAgent.Core.Models;
using OfficeAgent.ExcelAddIn.Dialogs;
using OfficeAgent.ExcelAddIn.Localization;

namespace OfficeAgent.ExcelAddIn
{
    public partial class AgentRibbon
    {
        private static string ProjectDropDownPlaceholderText => GetStrings().ProjectDropDownPlaceholderText;
        private const string DocumentationUrl = "https://github.com/SanChiaki/OfficeAgent";

        private readonly Dictionary<string, ProjectOption> projectOptionsByKey =
            new Dictionary<string, ProjectOption>(StringComparer.Ordinal);
        private readonly Dictionary<string, string> projectLabelsByKey =
            new Dictionary<string, string>(StringComparer.Ordinal);
        private readonly List<ProjectSelectorEntry> projectSelectorEntries =
            new List<ProjectSelectorEntry>();

        private bool isBoundToSyncController;
        private bool isBoundToTemplateController;
        private string lastControllerOwnedProjectDropDownText = HostLocalizedStrings.ForLocale("en").ProjectDropDownPlaceholderText;

        private sealed class ProjectSelectorEntry
        {
            public string Label { get; set; } = string.Empty;

            public string Tag { get; set; } = string.Empty;
        }

        private void AgentRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            ApplyLocalizedLabels();
            SetProjectDropDownText(ProjectDropDownPlaceholderText);
            RefreshTemplateButtonsFromController();
            BindToControllersAndRefresh();
        }

        private void ToggleTaskPaneButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.TaskPaneController?.Toggle();
        }

        private void LoginButton_Click(object sender, RibbonControlEventArgs e)
        {
            BeginLoginFlow(refreshProjectsAfterSuccess: true);
        }

        private void DocumentationButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                OpenUrlInDefaultBrowser(DocumentationUrl);
            }
            catch (Exception ex)
            {
                var strings = GetStrings();
                MessageBox.Show(
                    strings.DocumentationOpenFailedMessage(ex.Message),
                    strings.HostWindowTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void AboutButton_Click(object sender, RibbonControlEventArgs e)
        {
            var strings = GetStrings();
            MessageBox.Show(CreateAboutMessage(), strings.RibbonAboutDialogTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static void OpenUrlInDefaultBrowser(string url)
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = url,
                UseShellExecute = true,
            });
        }

        private static string CreateAboutMessage()
        {
            var assembly = typeof(AgentRibbon).Assembly;
            var strings = GetStrings();
            var assemblyVersion = assembly.GetName().Version?.ToString() ?? strings.UnknownText;

            return strings.AboutMessage(
                VersionInfo.AppVersion,
                assemblyVersion,
                GetBuildConfiguration(),
                GetAssemblyBuildTime(assembly));
        }

        private static string GetAssemblyBuildTime(Assembly assembly)
        {
            var location = assembly.Location;
            if (string.IsNullOrWhiteSpace(location) || !File.Exists(location))
            {
                return GetStrings().UnknownText;
            }

            return File.GetLastWriteTime(location).ToString("yyyy-MM-dd HH:mm:ss");
        }

        private static string GetBuildConfiguration()
        {
#if DEBUG
            return "Debug";
#else
            return "Release";
#endif
        }

        internal async void BeginLoginFlow(bool refreshProjectsAfterSuccess)
        {
            await ExecuteLoginFlow(refreshProjectsAfterSuccess).ConfigureAwait(true);
        }

        private async Task<bool> ExecuteLoginFlow(bool refreshProjectsAfterSuccess)
        {
            var settings = Globals.ThisAddIn.SettingsStore.Load();
            var ssoUrl = settings.SsoUrl;

            if (string.IsNullOrWhiteSpace(ssoUrl))
            {
                var strings = GetStrings();
                MessageBox.Show(strings.ConfigureSsoUrlFirstMessage, strings.HostWindowTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            loginButton.Label = GetStrings().RibbonLoginInProgressButtonLabel;
            loginButton.Enabled = false;

            try
            {
                using (var popup = new SsoLoginPopup(ssoUrl, settings.SsoLoginSuccessPath, Globals.ThisAddIn.SharedCookies, Globals.ThisAddIn.CookieStore))
                {
                    await popup.InitializeAsync().ConfigureAwait(true);
                    var dialogResult = popup.ShowDialog();
                    if (dialogResult != DialogResult.OK)
                    {
                        return false;
                    }
                }

                if (refreshProjectsAfterSuccess)
                {
                    PopulateProjectDropDown();
                    RefreshProjectDropDownFromController();
                }

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, GetStrings().HostWindowTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            finally
            {
                loginButton.Label = GetStrings().RibbonLoginButtonLabel;
                loginButton.Enabled = true;
            }
        }

        internal void RefreshProjectDropDownFromController()
        {
            if (!TryBindToSyncController())
            {
                return;
            }

            var syncController = Globals.ThisAddIn.RibbonSyncController;
            if (syncController == null)
            {
                return;
            }

            var noProjectRestoreText = GetNoProjectRestoreText(
                projectOptionsByKey.Count,
                syncController.ActiveProjectId,
                lastControllerOwnedProjectDropDownText);
            if (noProjectRestoreText != null)
            {
                SetProjectDropDownText(noProjectRestoreText);
                if (GetStrings().IsStickyProjectStatus(noProjectRestoreText))
                {
                    OfficeAgentLog.Warn(
                        "ribbon",
                        "project_selector.refresh_preserved_status",
                        $"Preserved project selector status. ProjectCount={projectOptionsByKey.Count}; Text={GetProjectDropDownDisplayText() ?? "<null>"}");
                }
                else
                {
                    OfficeAgentLog.Info(
                        "ribbon",
                        "project_selector.refresh_applied",
                        $"Refreshed project selector. ProjectCount={projectOptionsByKey.Count}; Text={GetProjectDropDownDisplayText() ?? "<null>"}");
                }

                return;
            }

            string text;
            if (!string.IsNullOrWhiteSpace(syncController.ActiveProjectId) &&
                !string.IsNullOrWhiteSpace(syncController.ActiveSystemKey))
            {
                var targetKey = ProjectSelectionKey.Build(syncController.ActiveSystemKey, syncController.ActiveProjectId);
                if (!projectLabelsByKey.TryGetValue(targetKey, out text))
                {
                    text = FormatProjectDropDownLabel(syncController.ActiveProjectId, syncController.ActiveProjectDisplayName);
                }
            }
            else
            {
                text = ProjectDropDownPlaceholderText;
            }

            if (string.IsNullOrWhiteSpace(text))
            {
                text = ProjectDropDownPlaceholderText;
            }

            SetProjectDropDownText(text);
            OfficeAgentLog.Info(
                "ribbon",
                "project_selector.refresh_applied",
                $"Refreshed project selector. ProjectCount={projectOptionsByKey.Count}; Text={GetProjectDropDownDisplayText() ?? "<null>"}");
        }

        private void PopulateProjectDropDown()
        {
            var syncController = Globals.ThisAddIn.RibbonSyncController;
            if (syncController == null)
            {
                return;
            }

            projectOptionsByKey.Clear();
            projectLabelsByKey.Clear();
            projectSelectorEntries.Clear();

            SetProjectDropDownText(ProjectDropDownPlaceholderText);

            try
            {
                var usedLabels = new HashSet<string>(StringComparer.Ordinal);
                var projects = syncController.GetProjects() ?? Array.Empty<ProjectOption>();
                foreach (var project in projects)
                {
                    var systemKey = project.SystemKey ?? string.Empty;
                    var projectId = project.ProjectId ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(systemKey) || string.IsNullOrWhiteSpace(projectId))
                    {
                        continue;
                    }

                    var projectKey = ProjectSelectionKey.Build(systemKey, projectId);
                    var projectLabel = CreateProjectDropDownLabel(project, usedLabels);
                    projectOptionsByKey[projectKey] = project;
                    projectLabelsByKey[projectKey] = projectLabel;
                    projectSelectorEntries.Add(new ProjectSelectorEntry
                    {
                        Label = projectLabel,
                        Tag = projectKey,
                    });
                }

                if (projectOptionsByKey.Count == 0)
                {
                    SetProjectDropDownStatus(GetStrings().ProjectDropDownNoAvailableProjectsText);
                    OfficeAgentLog.Warn("ribbon", "project_selector.empty", "Project list returned no available projects.");
                    ScheduleProjectLoadWarning(
                        GetStrings().ProjectListEmptyWarningMessage,
                        MessageBoxIcon.Warning);
                }
                else
                {
                    OfficeAgentLog.Info("ribbon", "project_selector.loaded", $"Loaded {projectOptionsByKey.Count} projects.");
                    OfficeAgentLog.Info(
                        "ribbon",
                        "project_selector.populate_applied",
                        $"Populated project selector. ProjectCount={projectOptionsByKey.Count}; Text={GetProjectDropDownDisplayText() ?? "<null>"}");
                }
            }
            catch (AuthenticationRequiredException ex)
            {
                SetProjectDropDownStatus(GetStrings().ProjectDropDownLoginRequiredText);
                OfficeAgentLog.Warn("ribbon", "project_selector.login_required", ex.Message);
                if (OperationResultDialog.ShowAuthenticationRequired(GetStrings().AuthenticationRequiredDefaultMessage))
                {
                    BeginLoginFlow(refreshProjectsAfterSuccess: true);
                }
            }
            catch (InvalidOperationException ex)
            {
                SetProjectDropDownStatus(GetStrings().ProjectDropDownLoadFailedText);
                OfficeAgentLog.Error("ribbon", "project_selector.load_failed", "Failed to load project list.", ex);
                ScheduleProjectLoadWarning(
                    GetStrings().ProjectListLoadFailedMessage(ex.Message),
                    MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                SetProjectDropDownStatus(GetStrings().ProjectDropDownLoadFailedText);
                OfficeAgentLog.Error("ribbon", "project_selector.load_failed", "Failed to load project list.", ex);
                ScheduleProjectLoadWarning(
                    GetStrings().ProjectListLoadFailedMessage(ex.Message),
                    MessageBoxIcon.Error);
            }
        }

        private void ScheduleProjectLoadWarning(string message, MessageBoxIcon icon)
        {
            var syncContext = SynchronizationContext.Current;
            OfficeAgentLog.Warn(
                "ribbon",
                "project_selector.warning_scheduled",
                $"Scheduling project selector warning. SynchronizationContext={syncContext?.GetType().FullName ?? "null"}; Message={message}");
            if (syncContext == null)
            {
                MessageBox.Show(message, GetStrings().HostWindowTitle, MessageBoxButtons.OK, icon);
                return;
            }

            syncContext.Post(
                _ => MessageBox.Show(message, GetStrings().HostWindowTitle, MessageBoxButtons.OK, icon),
                state: null);
        }

        private void SetProjectDropDownStatus(string label)
        {
            SetProjectDropDownText(label);
        }

        private void SetProjectDropDownText(string text)
        {
            var normalizedText = string.IsNullOrWhiteSpace(text)
                ? ProjectDropDownPlaceholderText
                : text;
            projectSelectorButton.Label = normalizedText;
            projectSelectorButton.ScreenTip = normalizedText;
            lastControllerOwnedProjectDropDownText = string.IsNullOrWhiteSpace(normalizedText)
                ? ProjectDropDownPlaceholderText
                : normalizedText;
            RibbonUI?.InvalidateControl(projectSelectorButton.Name);
        }

        private string CreateProjectDropDownLabel(ProjectOption project, ISet<string> usedLabels)
        {
            var baseLabel = FormatProjectDropDownLabel(project?.ProjectId ?? string.Empty, project?.DisplayName ?? string.Empty);
            var candidate = baseLabel;
            if (usedLabels.Contains(candidate))
            {
                candidate = $"{baseLabel} [{project?.SystemKey ?? string.Empty}/{project?.ProjectId ?? string.Empty}]";
            }

            usedLabels.Add(candidate);
            return candidate;
        }

        private static string FormatProjectDropDownLabel(string projectId, string displayName)
        {
            var normalizedProjectId = projectId?.Trim() ?? string.Empty;
            var normalizedDisplayName = displayName?.Trim() ?? string.Empty;

            if (string.IsNullOrWhiteSpace(normalizedProjectId))
            {
                return normalizedDisplayName;
            }

            if (string.IsNullOrWhiteSpace(normalizedDisplayName))
            {
                return normalizedProjectId;
            }

            return normalizedProjectId + "-" + normalizedDisplayName;
        }

        private void ProjectSelectorButton_Click(object sender, RibbonControlEventArgs e)
        {
            ShowProjectPickerDialog();
        }

        private void ShowProjectPickerDialog()
        {
            PopulateProjectDropDown();
            RefreshProjectDropDownFromController();

            if (projectOptionsByKey.Count == 0)
            {
                return;
            }

            var items = projectSelectorEntries
                .Where(item => projectOptionsByKey.ContainsKey(item.Tag))
                .Select(item => new ProjectPickerDialog.ProjectPickerItem(item.Label, projectOptionsByKey[item.Tag]))
                .ToArray();

            using (var dialog = new ProjectPickerDialog(items))
            {
                if (dialog.ShowDialog() != DialogResult.OK || dialog.SelectedProject == null)
                {
                    return;
                }

                Globals.ThisAddIn.RibbonSyncController?.SelectProject(dialog.SelectedProject);
            }
        }

        internal void BindToSyncControllerAndRefresh()
        {
            BindToControllersAndRefresh();
        }

        internal void BindToControllersAndRefresh()
        {
            ApplyLocalizedLabels();
            if (TryBindToSyncController())
            {
                Globals.ThisAddIn.RibbonSyncController?.RefreshActiveProjectFromSheetMetadata();
                RefreshProjectDropDownFromController();
            }

            if (TryBindToTemplateController())
            {
                Globals.ThisAddIn.RibbonTemplateController?.RefreshActiveTemplateStateFromSheetMetadata();
            }

            RefreshTemplateButtonsFromController();
        }

        private bool TryBindToSyncController()
        {
            if (isBoundToSyncController)
            {
                return true;
            }

            var syncController = Globals.ThisAddIn.RibbonSyncController;
            if (syncController == null)
            {
                return false;
            }

            syncController.ActiveProjectChanged += SyncController_ActiveProjectChanged;
            isBoundToSyncController = true;
            return true;
        }

        private bool TryBindToTemplateController()
        {
            if (isBoundToTemplateController)
            {
                return true;
            }

            var controller = Globals.ThisAddIn.RibbonTemplateController;
            if (controller == null)
            {
                return false;
            }

            controller.TemplateStateChanged += TemplateController_TemplateStateChanged;
            isBoundToTemplateController = true;
            return true;
        }

        internal void RefreshTemplateButtonsFromController()
        {
            var controller = Globals.ThisAddIn.RibbonTemplateController;
            applyTemplateButton.Enabled = controller?.CanApplyTemplate == true;
            saveTemplateButton.Enabled = controller?.CanSaveTemplate == true;
            saveAsTemplateButton.Enabled = controller?.CanSaveAsTemplate == true;
        }

        private string GetProjectDropDownDisplayText()
        {
            return projectSelectorButton.Label;
        }

        private void InitializeSheetButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.RibbonSyncController?.ExecuteInitializeCurrentSheet();
        }

        private void ApplyTemplateButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.RibbonTemplateController?.ExecuteApplyTemplate();
        }

        private void SaveTemplateButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.RibbonTemplateController?.ExecuteSaveTemplate();
        }

        private void SaveAsTemplateButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.RibbonTemplateController?.ExecuteSaveAsTemplate();
        }

        private void SyncController_ActiveProjectChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.RibbonTemplateController?.InvalidateRefreshState();
            Globals.ThisAddIn.RibbonTemplateController?.RefreshActiveTemplateStateFromSheetMetadata();
            RefreshProjectDropDownFromController();
            RefreshTemplateButtonsFromController();
        }

        private void TemplateController_TemplateStateChanged(object sender, EventArgs e)
        {
            RefreshTemplateButtonsFromController();
        }

        private static string GetNoProjectRestoreText(int projectOptionCount, string activeProjectId, string lastControllerOwnedText)
        {
            if (projectOptionCount != 0 || !string.IsNullOrWhiteSpace(activeProjectId))
            {
                return null;
            }

            return HostLocalizedStrings.IsKnownStickyProjectStatus(lastControllerOwnedText)
                ? lastControllerOwnedText
                : HostLocalizedStrings.ForLocale("en").ProjectDropDownPlaceholderText;
        }

        private void FullDownloadButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.RibbonSyncController?.ExecuteFullDownload();
        }

        private void PartialDownloadButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.RibbonSyncController?.ExecutePartialDownload();
        }

        private void FullUploadButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.RibbonSyncController?.ExecuteFullUpload();
        }

        private void PartialUploadButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.RibbonSyncController?.ExecutePartialUpload();
        }

        private void ApplyLocalizedLabels()
        {
            var strings = GetStrings();
            tab1.Label = strings.RibbonTabLabel;
            group1.Label = strings.RibbonAgentGroupLabel;
            toggleTaskPaneButton.Label = strings.RibbonAgentButtonLabel;
            groupProject.Label = strings.RibbonProjectGroupLabel;
            initializeSheetButton.Label = strings.RibbonInitializeSheetButtonLabel;
            groupTemplate.Label = strings.RibbonTemplateGroupLabel;
            applyTemplateButton.Label = strings.RibbonApplyTemplateButtonLabel;
            saveTemplateButton.Label = strings.RibbonSaveTemplateButtonLabel;
            saveAsTemplateButton.Label = strings.RibbonSaveAsTemplateButtonLabel;
            groupDataSync.Label = strings.RibbonDataSyncGroupLabel;
            fullDownloadButton.Label = strings.RibbonFullDownloadButtonLabel;
            partialDownloadButton.Label = strings.RibbonPartialDownloadButtonLabel;
            fullUploadButton.Label = strings.RibbonFullUploadButtonLabel;
            partialUploadButton.Label = strings.RibbonPartialUploadButtonLabel;
            group2.Label = strings.RibbonAccountGroupLabel;
            loginButton.Label = strings.RibbonLoginButtonLabel;
            groupHelp.Label = strings.RibbonHelpGroupLabel;
            documentationButton.Label = strings.RibbonDocumentationButtonLabel;
            aboutButton.Label = strings.RibbonAboutButtonLabel;
            projectSelectorButton.Label = ProjectDropDownPlaceholderText;
            projectSelectorButton.ScreenTip = ProjectDropDownPlaceholderText;
        }

        private static HostLocalizedStrings GetStrings()
        {
            var strings = Globals.ThisAddIn?.HostLocalizedStrings;
            return strings ?? HostLocalizedStrings.ForLocale("en");
        }
    }
}
