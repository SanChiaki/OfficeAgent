using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using OfficeAgent.Core;
using OfficeAgent.Core.Diagnostics;
using OfficeAgent.Core.Models;
using OfficeAgent.ExcelAddIn.Dialogs;

namespace OfficeAgent.ExcelAddIn
{
    public partial class AgentRibbon
    {
        private const string ProjectDropDownPlaceholderText = "先选择项目";
        private const string ProjectDropDownPlaceholderTag = "__no_project__";
        private const string SyntheticProjectDropDownTagPrefix = "__display__:";

        private static readonly string[] StickyNoProjectTexts =
        {
            "请先登录",
            "无可用项目",
            "项目加载失败",
        };

        private readonly Dictionary<string, ProjectOption> projectOptionsByKey =
            new Dictionary<string, ProjectOption>(StringComparer.Ordinal);
        private readonly Dictionary<string, ProjectOption> projectOptionsByLabel =
            new Dictionary<string, ProjectOption>(StringComparer.Ordinal);
        private readonly Dictionary<string, string> projectLabelsByKey =
            new Dictionary<string, string>(StringComparer.Ordinal);

        private bool isUpdatingProjectDropDown;
        private bool isBoundToSyncController;
        private bool isBoundToTemplateController;
        private string lastControllerOwnedProjectDropDownText = "先选择项目";

        private void AgentRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            SetProjectDropDownText("先选择项目");
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
                MessageBox.Show("\u8BF7\u5148\u5728\u8BBE\u7F6E\u4E2D\u914D\u7F6E SSO \u5730\u5740\u3002", "ISDP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            loginButton.Label = "\u767B\u5F55\u4E2D...";
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
                MessageBox.Show(ex.Message, "ISDP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            finally
            {
                loginButton.Label = "\u767B\u5F55";
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

            isUpdatingProjectDropDown = true;
            try
            {
                var noProjectRestoreText = GetNoProjectRestoreText(
                    projectOptionsByKey.Count,
                    syncController.ActiveProjectId,
                    lastControllerOwnedProjectDropDownText);
                if (noProjectRestoreText != null)
                {
                    SetProjectDropDownText(noProjectRestoreText);
                    if (IsStickyNoProjectText(noProjectRestoreText))
                    {
                        OfficeAgentLog.Warn(
                            "ribbon",
                            "project_dropdown.refresh_preserved_status",
                            $"Preserved project dropdown status. ItemCount={projectDropDown.Items.Count}; Text={GetProjectDropDownDisplayText() ?? "<null>"}");
                    }
                    else
                    {
                        OfficeAgentLog.Info(
                            "ribbon",
                            "project_dropdown.refresh_applied",
                            $"Refreshed project dropdown. ItemCount={projectDropDown.Items.Count}; Text={GetProjectDropDownDisplayText() ?? "<null>"}");
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
                    text = "先选择项目";
                }

                if (string.IsNullOrWhiteSpace(text))
                {
                    text = "先选择项目";
                }

                SetProjectDropDownText(text);
                OfficeAgentLog.Info(
                    "ribbon",
                    "project_dropdown.refresh_applied",
                    $"Refreshed project dropdown. ItemCount={projectDropDown.Items.Count}; Text={GetProjectDropDownDisplayText() ?? "<null>"}");
            }
            finally
            {
                isUpdatingProjectDropDown = false;
            }
        }

        private void PopulateProjectDropDown()
        {
            var syncController = Globals.ThisAddIn.RibbonSyncController;
            if (syncController == null)
            {
                return;
            }

            projectOptionsByKey.Clear();
            projectOptionsByLabel.Clear();
            projectLabelsByKey.Clear();

            isUpdatingProjectDropDown = true;
            try
            {
                projectDropDown.Items.Clear();
                AddProjectDropDownPlaceholderItem();
                SetProjectDropDownText("先选择项目");

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
                        projectOptionsByLabel[projectLabel] = project;
                        projectLabelsByKey[projectKey] = projectLabel;
                        projectDropDown.Items.Add(CreateProjectDropDownItem(projectLabel, projectKey));
                    }

                    if (projectOptionsByKey.Count == 0)
                    {
                        SetProjectDropDownStatus("无可用项目");
                        OfficeAgentLog.Warn("ribbon", "project_dropdown.empty", "Project list returned no available projects.");
                        ScheduleProjectLoadWarning(
                            "项目列表加载完成，但未获取到任何可用项目。\r\n请检查登录状态或项目接口返回。",
                            MessageBoxIcon.Warning);
                    }
                    else
                    {
                        OfficeAgentLog.Info("ribbon", "project_dropdown.loaded", $"Loaded {projectOptionsByKey.Count} projects.");
                        OfficeAgentLog.Info(
                            "ribbon",
                            "project_dropdown.populate_applied",
                            $"Populated project dropdown. ItemCount={projectDropDown.Items.Count}; Text={GetProjectDropDownDisplayText() ?? "<null>"}");
                    }
                }
                catch (AuthenticationRequiredException ex)
                {
                    SetProjectDropDownStatus("请先登录");
                    OfficeAgentLog.Warn("ribbon", "project_dropdown.login_required", ex.Message);
                    if (OperationResultDialog.ShowAuthenticationRequired(ex.Message))
                    {
                        BeginLoginFlow(refreshProjectsAfterSuccess: true);
                    }
                }
                catch (InvalidOperationException ex)
                {
                    SetProjectDropDownStatus("项目加载失败");
                    OfficeAgentLog.Error("ribbon", "project_dropdown.load_failed", "Failed to load project list.", ex);
                    ScheduleProjectLoadWarning(
                        $"项目列表加载失败。\r\n{ex.Message}",
                        MessageBoxIcon.Error);
                }
                catch (Exception ex)
                {
                    SetProjectDropDownStatus("项目加载失败");
                    OfficeAgentLog.Error("ribbon", "project_dropdown.load_failed", "Failed to load project list.", ex);
                    ScheduleProjectLoadWarning(
                        $"项目列表加载失败。\r\n{ex.Message}",
                        MessageBoxIcon.Error);
                }
            }
            finally
            {
                isUpdatingProjectDropDown = false;
            }
        }

        private void ScheduleProjectLoadWarning(string message, MessageBoxIcon icon)
        {
            var syncContext = SynchronizationContext.Current;
            OfficeAgentLog.Warn(
                "ribbon",
                "project_dropdown.warning_scheduled",
                $"Scheduling project dropdown warning. SynchronizationContext={syncContext?.GetType().FullName ?? "null"}; Message={message}");
            if (syncContext == null)
            {
                MessageBox.Show(message, "ISDP", MessageBoxButtons.OK, icon);
                return;
            }

            syncContext.Post(
                _ => MessageBox.Show(message, "ISDP", MessageBoxButtons.OK, icon),
                state: null);
        }

        private void SetProjectDropDownStatus(string label)
        {
            projectDropDown.Items.Clear();
            SetProjectDropDownText(label);
        }

        private void SetProjectDropDownText(string text)
        {
            var normalizedText = string.IsNullOrWhiteSpace(text)
                ? ProjectDropDownPlaceholderText
                : text;
            var selectedItem = EnsureProjectDropDownContainsDisplayItem(normalizedText);
            projectDropDown.SelectedItem = selectedItem;
            projectDropDown.Label = normalizedText;
            lastControllerOwnedProjectDropDownText = string.IsNullOrWhiteSpace(normalizedText)
                ? ProjectDropDownPlaceholderText
                : normalizedText;
            RibbonUI?.InvalidateControl(projectDropDown.Name);
        }

        private RibbonDropDownItem EnsureProjectDropDownContainsDisplayItem(string text)
        {
            var displayText = string.IsNullOrWhiteSpace(text)
                ? ProjectDropDownPlaceholderText
                : text;
            var existingItem = projectDropDown.Items
                .Cast<RibbonDropDownItem>()
                .FirstOrDefault(item => string.Equals(item?.Label, displayText, StringComparison.Ordinal));
            if (existingItem != null)
            {
                return existingItem;
            }

            var item = CreateProjectDropDownItem(displayText, BuildSyntheticProjectDropDownTag(displayText));
            projectDropDown.Items.Add(item);
            return item;
        }

        private static string BuildSyntheticProjectDropDownTag(string text)
        {
            return SyntheticProjectDropDownTagPrefix + (text ?? string.Empty);
        }

        private void AddProjectDropDownPlaceholderItem()
        {
            projectDropDown.Items.Add(CreateProjectDropDownItem(ProjectDropDownPlaceholderText, ProjectDropDownPlaceholderTag));
        }

        private RibbonDropDownItem CreateProjectDropDownItem(string label, string tag)
        {
            var item = Factory.CreateRibbonDropDownItem();
            item.Label = label ?? string.Empty;
            item.Tag = tag ?? string.Empty;
            return item;
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

        private void ProjectDropDown_ItemsLoading(object sender, RibbonControlEventArgs e)
        {
            PopulateProjectDropDown();
            RefreshProjectDropDownFromController();
        }

        private void RebuildProjectDropDownItemsFromCurrentState()
        {
            if (projectOptionsByKey.Count == 0)
            {
                return;
            }

            var items = projectDropDown.Items
                .Select(item => new
                {
                    item.Label,
                    Tag = item.Tag as string ?? string.Empty,
                })
                .Where(item => !string.Equals(item.Tag, ProjectDropDownPlaceholderTag, StringComparison.Ordinal))
                .ToArray();

            projectDropDown.Items.Clear();
            AddProjectDropDownPlaceholderItem();
            foreach (var item in items)
            {
                projectDropDown.Items.Add(CreateProjectDropDownItem(item.Label, item.Tag));
            }
        }

        private void ResetProjectDropDownItemsToPlaceholderOnly()
        {
            projectDropDown.Items.Clear();
            AddProjectDropDownPlaceholderItem();
        }

        private void ProjectDropDown_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            if (isUpdatingProjectDropDown)
            {
                return;
            }

            var selectedText = projectDropDown.SelectedItem?.Label ?? string.Empty;
            if (string.IsNullOrWhiteSpace(selectedText))
            {
                RestoreProjectDropDownFromController();
                return;
            }

            if (!projectOptionsByLabel.TryGetValue(selectedText, out var project))
            {
                RestoreProjectDropDownFromController();
                return;
            }

            Globals.ThisAddIn.RibbonSyncController?.SelectProject(project);
        }

        internal void BindToSyncControllerAndRefresh()
        {
            BindToControllersAndRefresh();
        }

        internal void BindToControllersAndRefresh()
        {
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
            return projectDropDown.SelectedItem?.Label ?? projectDropDown.Label;
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
            var syncController = Globals.ThisAddIn.RibbonSyncController;
            if (string.IsNullOrWhiteSpace(syncController?.ActiveProjectId))
            {
                ResetProjectDropDownItemsToPlaceholderOnly();
            }
            else
            {
                RebuildProjectDropDownItemsFromCurrentState();
            }

            Globals.ThisAddIn.RibbonTemplateController?.InvalidateRefreshState();
            Globals.ThisAddIn.RibbonTemplateController?.RefreshActiveTemplateStateFromSheetMetadata();
            RefreshProjectDropDownFromController();
            RefreshTemplateButtonsFromController();
        }

        private void TemplateController_TemplateStateChanged(object sender, EventArgs e)
        {
            RefreshTemplateButtonsFromController();
        }

        private void RestoreProjectDropDownFromController()
        {
            var syncController = Globals.ThisAddIn.RibbonSyncController;
            var noProjectRestoreText = GetNoProjectRestoreText(
                projectOptionsByKey.Count,
                syncController?.ActiveProjectId,
                lastControllerOwnedProjectDropDownText);
            if (noProjectRestoreText != null)
            {
                isUpdatingProjectDropDown = true;
                try
                {
                    SetProjectDropDownText(noProjectRestoreText);
                }
                finally
                {
                    isUpdatingProjectDropDown = false;
                }

                return;
            }

            RefreshProjectDropDownFromController();
        }

        private static string GetNoProjectRestoreText(int projectOptionCount, string activeProjectId, string lastControllerOwnedText)
        {
            if (projectOptionCount != 0 || !string.IsNullOrWhiteSpace(activeProjectId))
            {
                return null;
            }

            return IsStickyNoProjectText(lastControllerOwnedText)
                ? lastControllerOwnedText
                : "先选择项目";
        }

        private static bool IsStickyNoProjectText(string text)
        {
            return Array.IndexOf(StickyNoProjectTexts, text ?? string.Empty) >= 0;
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
    }
}
