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

            Assert.Contains("if (selected == null && projectOptionsByKey.Count == 0)", ribbonCodeText, StringComparison.Ordinal);
        }

        [Fact]
        public void PopulateProjectDropDownSelectsPlaceholderItemBeforeAnyProjectIsChosen()
        {
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            Assert.Contains("var placeholderItem = CreateProjectDropDownItem(\"先选择项目\", EmptyProjectTag);", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("SelectProjectDropDownItem(placeholderItem);", ribbonCodeText, StringComparison.Ordinal);
        }

        [Fact]
        public void ProjectDropDownDisplaysItemTextInsteadOfSeparateControlCaption()
        {
            var designerText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.Designer.cs"));

            Assert.Contains("this.projectDropDown.ShowItemLabel = true;", designerText, StringComparison.Ordinal);
            Assert.Contains("this.projectDropDown.ShowLabel = false;", designerText, StringComparison.Ordinal);
        }

        [Fact]
        public void ProjectDropDownSelectionSetsSelectedItemIndexForOfficeHostCompatibility()
        {
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            Assert.Contains("projectDropDown.SelectedItemIndex = projectDropDown.Items.IndexOf(item);", ribbonCodeText, StringComparison.Ordinal);
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
