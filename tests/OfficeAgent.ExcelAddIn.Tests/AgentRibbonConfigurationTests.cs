using System;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class AgentRibbonConfigurationTests
    {
        [Fact]
        public void TaskPaneButtonDoesNotDependOnRuntimeImageAssignment()
        {
            var designerText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.Designer.cs"));
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            Assert.Contains("this.toggleTaskPaneButton.ShowImage = false;", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("RibbonControlSize.RibbonControlSizeLarge", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("toggleTaskPaneButton.Image = Properties.Resources.Logo;", ribbonCodeText, StringComparison.Ordinal);
        }

        [Fact]
        public void TaskPaneGroupUsesStableDedicatedRibbonIdentifiers()
        {
            var designerText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.Designer.cs"));

            Assert.Contains("this.group1.Name = \"groupAgent\";", designerText, StringComparison.Ordinal);
            Assert.Contains("this.toggleTaskPaneButton.Name = \"openTaskPaneButton\";", designerText, StringComparison.Ordinal);
        }

        [Fact]
        public void RibbonUsesDedicatedCustomTabInsteadOfBuiltInAddInsTab()
        {
            var designerText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.Designer.cs"));

            Assert.DoesNotContain("this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("this.tab1.ControlId.OfficeId = \"TabAddIns\";", designerText, StringComparison.Ordinal);
        }

        [Fact]
        public void TaskPaneGroupIsInsertedBeforeProjectGroup()
        {
            var designerText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.Designer.cs"));

            var taskPaneGroupIndex = designerText.IndexOf("this.tab1.Groups.Add(this.group1);", StringComparison.Ordinal);
            var projectGroupIndex = designerText.IndexOf("this.tab1.Groups.Add(this.groupProject);", StringComparison.Ordinal);

            Assert.True(taskPaneGroupIndex >= 0);
            Assert.True(projectGroupIndex > taskPaneGroupIndex);
        }

        [Fact]
        public void LoginRefreshesProjectListAfterPopupCloses()
        {
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            var showDialogIndex = ribbonCodeText.IndexOf("popup.ShowDialog();", StringComparison.Ordinal);
            var repopulateIndex = ribbonCodeText.IndexOf("PopulateProjectDropDown();", showDialogIndex, StringComparison.Ordinal);

            Assert.True(showDialogIndex >= 0);
            Assert.True(repopulateIndex > showDialogIndex);
        }

        [Fact]
        public void ProjectLoadingWarnsUserWhenAuthenticationIsRequired()
        {
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            Assert.Contains("catch (InvalidOperationException ex)", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("请先登录", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("MessageBoxIcon.Warning", ribbonCodeText, StringComparison.Ordinal);
        }

        [Fact]
        public void ProjectLoadingMarksDropdownAsLoginRequiredWhenAuthenticationFails()
        {
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            Assert.Contains("SetProjectDropDownStatus(\"请先登录\")", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("ScheduleProjectLoadWarning(", ribbonCodeText, StringComparison.Ordinal);
        }

        [Fact]
        public void EmptyProjectListsWarnUserInsteadOfStayingSilentlyEmpty()
        {
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            Assert.Contains("if (projectOptionsByKey.Count == 0)", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("未获取到任何可用项目", ribbonCodeText, StringComparison.Ordinal);
        }

        [Fact]
        public void ProjectLoadingUsesDedicatedStatusItemsInsteadOfRibbonLabelOnly()
        {
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            Assert.Contains("SetProjectDropDownStatus(\"请先登录\")", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("SetProjectDropDownStatus(\"无可用项目\")", ribbonCodeText, StringComparison.Ordinal);
        }

        [Fact]
        public void RefreshProjectDropDownPreservesStatusWhenNoProjectsAreAvailable()
        {
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            Assert.Contains("if (projectOptionsByKey.Count == 0 && string.IsNullOrWhiteSpace(syncController.ActiveProjectId))", ribbonCodeText, StringComparison.Ordinal);
        }

        [Fact]
        public void PopulateProjectDropDownSetsPlaceholderTextBeforeAnyProjectIsChosen()
        {
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            Assert.Contains("SetProjectDropDownText(\"先选择项目\");", ribbonCodeText, StringComparison.Ordinal);
        }

        [Fact]
        public void ProjectDropDownDisplaysItemTextInsteadOfSeparateControlCaption()
        {
            var designerText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.Designer.cs"));

            Assert.Contains("this.projectDropDown.ShowLabel = false;", designerText, StringComparison.Ordinal);
        }

        [Fact]
        public void ProjectSelectorUsesTextValueForOfficeHostCompatibility()
        {
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            Assert.Contains("projectDropDown.Text = text ?? string.Empty;", ribbonCodeText, StringComparison.Ordinal);
        }

        [Fact]
        public void ProjectSelectorUsesComboBoxItemsLoadingToRefreshProjectsOnOpen()
        {
            var designerText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.Designer.cs"));
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            Assert.Contains(
                "this.projectDropDown = Factory.CreateRibbonComboBox();",
                designerText,
                StringComparison.Ordinal);
            Assert.Contains(
                "this.projectDropDown.ItemsLoading += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ProjectDropDown_ItemsLoading);",
                designerText,
                StringComparison.Ordinal);
            Assert.Contains(
                "this.projectDropDown.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ProjectDropDown_TextChanged);",
                designerText,
                StringComparison.Ordinal);
            Assert.DoesNotContain("this.projectDropDown.ButtonClick +=", designerText, StringComparison.Ordinal);
            Assert.Contains("private void ProjectDropDown_ItemsLoading(object sender, RibbonControlEventArgs e)", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("private void ProjectDropDown_TextChanged(object sender, RibbonControlEventArgs e)", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("PopulateProjectDropDown();", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("RefreshProjectDropDownFromController();", ribbonCodeText, StringComparison.Ordinal);
        }

        [Fact]
        public void RibbonLoadDoesNotPreloadProjectListBeforeUserOpensSelector()
        {
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            var methodStart = ribbonCodeText.IndexOf("private void AgentRibbon_Load(object sender, RibbonUIEventArgs e)", StringComparison.Ordinal);
            var nextMethodStart = ribbonCodeText.IndexOf("private void ToggleTaskPaneButton_Click", methodStart, StringComparison.Ordinal);

            Assert.True(methodStart >= 0);
            Assert.True(nextMethodStart > methodStart);

            var loadMethodText = ribbonCodeText.Substring(methodStart, nextMethodStart - methodStart);
            Assert.DoesNotContain("PopulateProjectDropDown();", loadMethodText, StringComparison.Ordinal);
            Assert.Contains("syncController.RefreshActiveProjectFromSheetMetadata();", loadMethodText, StringComparison.Ordinal);
            Assert.Contains("RefreshProjectDropDownFromController();", loadMethodText, StringComparison.Ordinal);
        }

        private static string ResolveRepositoryPath(params string[] segments)
        {
            return Path.GetFullPath(Path.Combine(new[]
            {
                AppContext.BaseDirectory,
                "..",
                "..",
                "..",
                "..",
                "..",
            }.Concat(segments).ToArray()));
        }
    }
}
