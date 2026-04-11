using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using OfficeAgent.Core.Models;

namespace OfficeAgent.ExcelAddIn
{
    public partial class AgentRibbon
    {
        private const string EmptyProjectTag = "__empty__";

        private readonly Dictionary<string, ProjectOption> projectOptionsById =
            new Dictionary<string, ProjectOption>(StringComparer.Ordinal);

        private bool isUpdatingProjectDropDown;

        private void AgentRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            try
            {
                toggleTaskPaneButton.Image = Properties.Resources.Logo;
            }
            catch
            {
                // Logo is optional; skip if not found
            }

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
                popup.ShowDialog();
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

            var targetTag = string.IsNullOrWhiteSpace(syncController.ActiveProjectId)
                ? EmptyProjectTag
                : syncController.ActiveProjectId;

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
                        syncController.ActiveProjectId);
                    projectDropDown.Items.Add(selected);
                }

                if (selected == null)
                {
                    selected = projectDropDown.Items
                        .OfType<RibbonDropDownItem>()
                        .FirstOrDefault(item => string.Equals(item.Tag as string, EmptyProjectTag, StringComparison.Ordinal));
                }

                projectDropDown.SelectedItem = selected;
                projectDropDown.Label = syncController.ActiveProjectDisplayName;
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

            projectOptionsById.Clear();

            isUpdatingProjectDropDown = true;
            try
            {
                projectDropDown.Items.Clear();
                projectDropDown.Items.Add(CreateProjectDropDownItem("先选择项目", EmptyProjectTag));

                var projects = syncController.GetProjects() ?? Array.Empty<ProjectOption>();
                foreach (var project in projects)
                {
                    var projectId = project.ProjectId ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(projectId))
                    {
                        continue;
                    }

                    projectOptionsById[projectId] = project;
                    projectDropDown.Items.Add(CreateProjectDropDownItem(project.DisplayName, projectId));
                }
            }
            finally
            {
                isUpdatingProjectDropDown = false;
            }
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

            if (!projectOptionsById.TryGetValue(selectedTag, out var project))
            {
                return;
            }

            Globals.ThisAddIn.RibbonSyncController?.SelectProject(project);
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

        private void IncrementalUploadButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.RibbonSyncController?.ExecuteIncrementalUpload();
        }
    }
}
