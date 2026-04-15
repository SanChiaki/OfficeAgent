using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using OfficeAgent.Core.Diagnostics;
using OfficeAgent.Core.Models;

namespace OfficeAgent.ExcelAddIn
{
    public partial class AgentRibbon
    {
        private const string EmptyProjectTag = "__empty__";
        private const string StatusProjectTag = "__status__";

        private readonly Dictionary<string, ProjectOption> projectOptionsByKey =
            new Dictionary<string, ProjectOption>(StringComparer.Ordinal);

        private bool isUpdatingProjectDropDown;

        private void AgentRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            var syncController = Globals.ThisAddIn.RibbonSyncController;
            if (syncController == null)
            {
                return;
            }

            syncController.ActiveProjectChanged += SyncController_ActiveProjectChanged;
            PopulateProjectDropDown();
            syncController.RefreshActiveProjectFromSheetMetadata();
            RefreshProjectDropDownFromController();
        }

        private void ToggleTaskPaneButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.TaskPaneController?.Toggle();
        }

        private async void LoginButton_Click(object sender, RibbonControlEventArgs e)
        {
            var settings = Globals.ThisAddIn.SettingsStore.Load();
            var ssoUrl = settings.SsoUrl;

            if (string.IsNullOrWhiteSpace(ssoUrl))
            {
                MessageBox.Show("\u8BF7\u5148\u5728\u8BBE\u7F6E\u4E2D\u914D\u7F6E SSO \u5730\u5740\u3002", "Resy AI", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            loginButton.Label = "\u767B\u5F55\u4E2D...";
            loginButton.Enabled = false;

            try
            {
                var popup = new SsoLoginPopup(ssoUrl, settings.SsoLoginSuccessPath, Globals.ThisAddIn.SharedCookies, Globals.ThisAddIn.CookieStore);
                await popup.InitializeAsync();
                var dialogResult = popup.ShowDialog();
                if (dialogResult == DialogResult.OK)
                {
                    PopulateProjectDropDown();
                    RefreshProjectDropDownFromController();
                }
            }
            finally
            {
                loginButton.Label = "\u767B\u5F55";
                loginButton.Enabled = true;
            }
        }

        internal void RefreshProjectDropDownFromController()
        {
            var syncController = Globals.ThisAddIn.RibbonSyncController;
            if (syncController == null)
            {
                return;
            }

            var targetTag = string.IsNullOrWhiteSpace(syncController.ActiveProjectId) ||
                            string.IsNullOrWhiteSpace(syncController.ActiveSystemKey)
                ? EmptyProjectTag
                : ProjectSelectionKey.Build(syncController.ActiveSystemKey, syncController.ActiveProjectId);

            isUpdatingProjectDropDown = true;
            try
            {
                var selected = projectDropDown.Items
                    .OfType<RibbonDropDownItem>()
                    .FirstOrDefault(item => string.Equals(item.Tag as string, targetTag, StringComparison.Ordinal));

                if (selected == null && !string.IsNullOrWhiteSpace(syncController.ActiveProjectId))
                {
                    selected = CreateProjectDropDownItem(
                        syncController.ActiveProjectDisplayName,
                        ProjectSelectionKey.Build(syncController.ActiveSystemKey, syncController.ActiveProjectId));
                    projectDropDown.Items.Add(selected);
                }

                if (selected == null && projectOptionsByKey.Count == 0)
                {
                    OfficeAgentLog.Warn("ribbon", "project_dropdown.refresh_preserved_status", "Skipped project dropdown refresh because no projects are currently available.");
                    return;
                }

                if (selected == null)
                {
                    selected = projectDropDown.Items
                        .OfType<RibbonDropDownItem>()
                        .FirstOrDefault(item => string.Equals(item.Tag as string, EmptyProjectTag, StringComparison.Ordinal));
                }

                SelectProjectDropDownItem(selected);
                projectDropDown.Label = syncController.ActiveProjectDisplayName;
                OfficeAgentLog.Info(
                    "ribbon",
                    "project_dropdown.refresh_applied",
                    $"Refreshed project dropdown. ItemCount={projectDropDown.Items.Count}; Selected={(projectDropDown.SelectedItem?.Label ?? "<null>")}; SelectedIndex={projectDropDown.SelectedItemIndex}");
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

            isUpdatingProjectDropDown = true;
            try
            {
                projectDropDown.Items.Clear();
                var placeholderItem = CreateProjectDropDownItem("先选择项目", EmptyProjectTag);
                projectDropDown.Items.Add(placeholderItem);
                SelectProjectDropDownItem(placeholderItem);

                try
                {
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
                        projectOptionsByKey[projectKey] = project;
                        projectDropDown.Items.Add(CreateProjectDropDownItem(project.DisplayName, projectKey));
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
                            $"Populated project dropdown. ItemCount={projectDropDown.Items.Count}; Selected={(projectDropDown.SelectedItem?.Label ?? "<null>")}; SelectedIndex={projectDropDown.SelectedItemIndex}");
                    }
                }
                catch (InvalidOperationException ex)
                {
                    MessageBoxIcon icon;
                    if (ex.Message.IndexOf("请先登录", StringComparison.Ordinal) >= 0)
                    {
                        icon = MessageBoxIcon.Warning;
                        SetProjectDropDownStatus("请先登录");
                        OfficeAgentLog.Warn("ribbon", "project_dropdown.login_required", ex.Message);
                    }
                    else
                    {
                        icon = MessageBoxIcon.Error;
                        SetProjectDropDownStatus("项目加载失败");
                        OfficeAgentLog.Error("ribbon", "project_dropdown.load_failed", "Failed to load project list.", ex);
                    }

                    ScheduleProjectLoadWarning(
                        $"项目列表加载失败。\r\n{ex.Message}",
                        icon);
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
                MessageBox.Show(message, "Resy AI", MessageBoxButtons.OK, icon);
                return;
            }

            syncContext.Post(
                _ => MessageBox.Show(message, "Resy AI", MessageBoxButtons.OK, icon),
                state: null);
        }

        private void SetProjectDropDownStatus(string label)
        {
            projectDropDown.Items.Clear();
            var statusItem = CreateProjectDropDownItem(label, StatusProjectTag);
            projectDropDown.Items.Add(statusItem);
            SelectProjectDropDownItem(statusItem);
            projectDropDown.Label = label ?? string.Empty;
        }

        private void SelectProjectDropDownItem(RibbonDropDownItem item)
        {
            projectDropDown.SelectedItem = item;
            projectDropDown.SelectedItemIndex = projectDropDown.Items.IndexOf(item);
        }

        private RibbonDropDownItem CreateProjectDropDownItem(string label, string tag)
        {
            var item = Factory.CreateRibbonDropDownItem();
            item.Label = label ?? string.Empty;
            item.Tag = tag ?? string.Empty;
            return item;
        }

        private void ProjectDropDown_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            if (isUpdatingProjectDropDown)
            {
                return;
            }

            var selectedTag = projectDropDown.SelectedItem?.Tag as string ?? string.Empty;
            if (string.IsNullOrWhiteSpace(selectedTag) || string.Equals(selectedTag, EmptyProjectTag, StringComparison.Ordinal))
            {
                return;
            }

            if (!projectOptionsByKey.TryGetValue(selectedTag, out var project))
            {
                return;
            }

            Globals.ThisAddIn.RibbonSyncController?.SelectProject(project);
        }

        private void InitializeSheetButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.RibbonSyncController?.ExecuteInitializeCurrentSheet();
        }

        private void SyncController_ActiveProjectChanged(object sender, EventArgs e)
        {
            RefreshProjectDropDownFromController();
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
