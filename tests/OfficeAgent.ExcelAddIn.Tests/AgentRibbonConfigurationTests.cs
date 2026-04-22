using System;
using System.IO;
using System.Linq;
using System.Reflection;
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
        public void RibbonTabStaysIsdpWhileAgentGroupUsesIsdpAi()
        {
            var designerText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.Designer.cs"));
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            Assert.Contains("this.tab1.Label = \"ISDP\";", designerText, StringComparison.Ordinal);
            Assert.Contains("this.group1.Label = \"ISDP AI\";", designerText, StringComparison.Ordinal);
            Assert.Contains("this.toggleTaskPaneButton.Label = \"ISDP AI\";", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("Resy AI", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("Resy AI", ribbonCodeText, StringComparison.Ordinal);
        }

        [Fact]
        public void UploadGroupOmitsFullUploadButton()
        {
            var designerText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.Designer.cs"));

            Assert.DoesNotContain("this.groupUpload.Items.Add(this.fullUploadButton);", designerText, StringComparison.Ordinal);
            Assert.Contains("this.groupUpload.Items.Add(this.partialUploadButton);", designerText, StringComparison.Ordinal);
        }

        [Fact]
        public void DownloadGroupOmitsFullDownloadButton()
        {
            var designerText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.Designer.cs"));

            Assert.DoesNotContain("this.groupDownload.Items.Add(this.fullDownloadButton);", designerText, StringComparison.Ordinal);
            Assert.Contains("this.groupDownload.Items.Add(this.partialDownloadButton);", designerText, StringComparison.Ordinal);
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
        public void TemplateGroupAppearsAfterProjectGroupAndBeforeDownloadGroup()
        {
            var designerText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.Designer.cs"));

            var projectGroupIndex = designerText.IndexOf("this.tab1.Groups.Add(this.groupProject);", StringComparison.Ordinal);
            var templateGroupIndex = designerText.IndexOf("this.tab1.Groups.Add(this.groupTemplate);", StringComparison.Ordinal);
            var downloadGroupIndex = designerText.IndexOf("this.tab1.Groups.Add(this.groupDownload);", StringComparison.Ordinal);

            Assert.True(projectGroupIndex >= 0);
            Assert.True(templateGroupIndex > projectGroupIndex);
            Assert.True(downloadGroupIndex > templateGroupIndex);
        }

        [Fact]
        public void TemplateGroupContainsApplySaveAndSaveAsButtons()
        {
            var designerText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.Designer.cs"));

            Assert.Contains("this.groupTemplate.Items.Add(this.applyTemplateButton);", designerText, StringComparison.Ordinal);
            Assert.Contains("this.groupTemplate.Items.Add(this.saveTemplateButton);", designerText, StringComparison.Ordinal);
            Assert.Contains("this.groupTemplate.Items.Add(this.saveAsTemplateButton);", designerText, StringComparison.Ordinal);
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

            Assert.Contains("catch (AuthenticationRequiredException ex)", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("ShowAuthenticationRequired", ribbonCodeText, StringComparison.Ordinal);
        }

        [Fact]
        public void ProjectLoadingMarksDropdownAsLoginRequiredWhenAuthenticationFails()
        {
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            Assert.Contains("SetProjectDropDownStatus(\"请先登录\")", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("ExecuteLoginFlow", ribbonCodeText, StringComparison.Ordinal);
        }

        [Fact]
        public void AuthenticationPromptOffersPointMeToLoginButton()
        {
            var dialogCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "Dialogs",
                "OperationResultDialog.cs"));

            Assert.Contains("点我登录", dialogCodeText, StringComparison.Ordinal);
            Assert.Contains("ShowAuthenticationRequired", dialogCodeText, StringComparison.Ordinal);
        }

        [Fact]
        public void AuthenticationPromptSizesButtonsFromMeasuredTextInsteadOfFixedWidths()
        {
            var dialogCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "Dialogs",
                "OperationResultDialog.cs"));

            Assert.Contains("TextRenderer.MeasureText", dialogCodeText, StringComparison.Ordinal);
            Assert.DoesNotContain("new Rectangle(154, 88, 90, 28)", dialogCodeText, StringComparison.Ordinal);
            Assert.DoesNotContain("new Rectangle(250, 88, 90, 28)", dialogCodeText, StringComparison.Ordinal);
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
        public void RefreshProjectDropDownUsesNoProjectRestoreTextWhenNoProjectsAreAvailable()
        {
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            Assert.Contains("var noProjectRestoreText = GetNoProjectRestoreText(", ribbonCodeText, StringComparison.Ordinal);
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
        public void PopulateProjectDropDownAddsPlaceholderItemBeforeLoadedProjects()
        {
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            Assert.Contains("AddProjectDropDownPlaceholderItem();", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("projectDropDown.Items.Add(CreateProjectDropDownItem(ProjectDropDownPlaceholderText, ProjectDropDownPlaceholderTag));", ribbonCodeText, StringComparison.Ordinal);
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
        public void ProjectDropDownUsesWideSizingStringToExpandProjectRibbonGroup()
        {
            var designerText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.Designer.cs"));

            Assert.Contains("this.projectDropDown.SizeString =", designerText, StringComparison.Ordinal);
        }

        [Fact]
        public void ProjectSelectorUsesSelectedItemForOfficeHostCompatibility()
        {
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            Assert.Contains("projectDropDown.SelectedItem = selectedItem;", ribbonCodeText, StringComparison.Ordinal);
        }

        [Fact]
        public void ProjectSelectorInvalidatesRibbonControlAfterProgrammaticSelectionChanges()
        {
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            Assert.Contains("RibbonUI?.InvalidateControl(projectDropDown.Name);", ribbonCodeText, StringComparison.Ordinal);
        }

        [Fact]
        public void ProjectSelectorEnsuresDropDownContainsDisplayItemBeforeSelectingIt()
        {
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            var methodStart = ribbonCodeText.IndexOf("private void SetProjectDropDownText(string text)", StringComparison.Ordinal);
            var nextMethodStart = ribbonCodeText.IndexOf("private void AddProjectDropDownPlaceholderItem()", methodStart, StringComparison.Ordinal);

            Assert.True(methodStart >= 0);
            Assert.True(nextMethodStart > methodStart);

            var methodBody = ribbonCodeText.Substring(methodStart, nextMethodStart - methodStart);
            Assert.Contains("var selectedItem = EnsureProjectDropDownContainsDisplayItem(normalizedText);", methodBody, StringComparison.Ordinal);
            Assert.Contains("projectDropDown.SelectedItem = selectedItem;", methodBody, StringComparison.Ordinal);
        }

        [Fact]
        public void ProjectSelectorDefinesHelperToAddSyntheticDisplayItemWhenCurrentLabelIsMissing()
        {
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            Assert.Contains("private RibbonDropDownItem EnsureProjectDropDownContainsDisplayItem(string text)", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("var item = CreateProjectDropDownItem(displayText, BuildSyntheticProjectDropDownTag(displayText));", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("projectDropDown.Items.Add(item);", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("return item;", ribbonCodeText, StringComparison.Ordinal);
        }

        [Fact]
        public void ProjectSelectorUsesDropDownItemsLoadingToRefreshProjectsOnOpen()
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
                "this.projectDropDown = Factory.CreateRibbonDropDown();",
                designerText,
                StringComparison.Ordinal);
            Assert.Contains(
                "this.projectDropDown.ItemsLoading += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ProjectDropDown_ItemsLoading);",
                designerText,
                StringComparison.Ordinal);
            Assert.Contains(
                "this.projectDropDown.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ProjectDropDown_SelectionChanged);",
                designerText,
                StringComparison.Ordinal);
            Assert.DoesNotContain("this.projectDropDown.ButtonClick +=", designerText, StringComparison.Ordinal);
            Assert.Contains("private void ProjectDropDown_ItemsLoading(object sender, RibbonControlEventArgs e)", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("private void ProjectDropDown_SelectionChanged(object sender, RibbonControlEventArgs e)", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("PopulateProjectDropDown();", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("RefreshProjectDropDownFromController();", ribbonCodeText, StringComparison.Ordinal);
        }

        [Fact]
        public void ActiveProjectChangeRebuildsExistingDropdownItemsBeforeRefreshingText()
        {
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            var methodStart = ribbonCodeText.IndexOf("private void SyncController_ActiveProjectChanged(object sender, EventArgs e)", StringComparison.Ordinal);
            var nextMethodStart = ribbonCodeText.IndexOf("private void RestoreProjectDropDownFromController()", methodStart, StringComparison.Ordinal);

            Assert.True(methodStart >= 0);
            Assert.True(nextMethodStart > methodStart);

            var methodBody = ribbonCodeText.Substring(methodStart, nextMethodStart - methodStart);
            Assert.Contains("RebuildProjectDropDownItemsFromCurrentState();", methodBody, StringComparison.Ordinal);
            Assert.Contains("RefreshProjectDropDownFromController();", methodBody, StringComparison.Ordinal);
        }

        [Fact]
        public void ActiveProjectChangeWithoutBoundProjectResetsDropdownItemsToPlaceholderOnly()
        {
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            var methodStart = ribbonCodeText.IndexOf("private void SyncController_ActiveProjectChanged(object sender, EventArgs e)", StringComparison.Ordinal);
            var nextMethodStart = ribbonCodeText.IndexOf("private void RestoreProjectDropDownFromController()", methodStart, StringComparison.Ordinal);

            Assert.True(methodStart >= 0);
            Assert.True(nextMethodStart > methodStart);

            var methodBody = ribbonCodeText.Substring(methodStart, nextMethodStart - methodStart);
            Assert.Contains("string.IsNullOrWhiteSpace(syncController?.ActiveProjectId)", methodBody, StringComparison.Ordinal);
            Assert.Contains("ResetProjectDropDownItemsToPlaceholderOnly();", methodBody, StringComparison.Ordinal);
        }

        [Fact]
        public void DropdownItemRebuildClearsAndReaddsExistingItemsWithoutReloadingProjects()
        {
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            Assert.Contains("private void RebuildProjectDropDownItemsFromCurrentState()", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("projectDropDown.Items.Clear();", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("AddProjectDropDownPlaceholderItem();", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("projectDropDown.Items.Add(CreateProjectDropDownItem(item.Label, item.Tag));", ribbonCodeText, StringComparison.Ordinal);
        }

        [Fact]
        public void PlaceholderResetClearsProjectItemsAndKeepsOnlyPlaceholder()
        {
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            Assert.Contains("private void ResetProjectDropDownItemsToPlaceholderOnly()", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("projectDropDown.Items.Clear();", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("AddProjectDropDownPlaceholderItem();", ribbonCodeText, StringComparison.Ordinal);
        }

        [Fact]
        public void ProjectSelectionLeavesDropdownResetToControllerRefreshFlow()
        {
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            var methodStart = ribbonCodeText.IndexOf("private void ProjectDropDown_SelectionChanged(object sender, RibbonControlEventArgs e)", StringComparison.Ordinal);
            var nextMethodStart = ribbonCodeText.IndexOf("internal void BindToSyncControllerAndRefresh()", methodStart, StringComparison.Ordinal);

            Assert.True(methodStart >= 0);
            Assert.True(nextMethodStart > methodStart);

            var methodBody = ribbonCodeText.Substring(methodStart, nextMethodStart - methodStart);
            Assert.Contains("Globals.ThisAddIn.RibbonSyncController?.SelectProject(project);", methodBody, StringComparison.Ordinal);
            Assert.DoesNotContain("SetProjectDropDownText(", methodBody, StringComparison.Ordinal);
            Assert.DoesNotContain("RefreshProjectDropDownFromController();", methodBody, StringComparison.Ordinal);
        }

        [Fact]
        public void ProjectSelectionRestoresControllerDisplayForMissingOrUnknownSelectionViaWrapper()
        {
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            var methodStart = ribbonCodeText.IndexOf("private void ProjectDropDown_SelectionChanged(object sender, RibbonControlEventArgs e)", StringComparison.Ordinal);
            var nextMethodStart = ribbonCodeText.IndexOf("internal void BindToSyncControllerAndRefresh()", methodStart, StringComparison.Ordinal);

            Assert.True(methodStart >= 0);
            Assert.True(nextMethodStart > methodStart);

            var methodBody = ribbonCodeText.Substring(methodStart, nextMethodStart - methodStart);
            Assert.Contains("RestoreProjectDropDownFromController();", methodBody, StringComparison.Ordinal);
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
            Assert.Contains("BindToControllersAndRefresh();", loadMethodText, StringComparison.Ordinal);
        }

        [Fact]
        public void RibbonDefinesLazyControllerBindingHelperForStartupOrdering()
        {
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            Assert.Contains("private bool TryBindToSyncController()", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("syncController.ActiveProjectChanged += SyncController_ActiveProjectChanged;", ribbonCodeText, StringComparison.Ordinal);
        }

        [Fact]
        public void RibbonBindsToTemplateControllerAndRefreshesTemplateButtons()
        {
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            Assert.Contains("BindToControllersAndRefresh()", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("TryBindToTemplateController()", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("RefreshTemplateButtonsFromController();", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("ApplyTemplateButton_Click", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("SaveTemplateButton_Click", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("SaveAsTemplateButton_Click", ribbonCodeText, StringComparison.Ordinal);
        }

        [Fact]
        public void ThisAddInInvalidatesSettingsCacheWhenSettingsSheetChanges()
        {
            var addInCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "ThisAddIn.cs"));

            Assert.Contains("Application.SheetChange += Application_SheetChange;", addInCodeText, StringComparison.Ordinal);
            Assert.Contains("private void Application_SheetChange(object sh, ExcelInterop.Range target)", addInCodeText, StringComparison.Ordinal);
            Assert.Contains("string.Equals(sheetName, \"ISDP_Setting\", StringComparison.OrdinalIgnoreCase)", addInCodeText, StringComparison.Ordinal);
            Assert.Contains("metadataStore.InvalidateCache();", addInCodeText, StringComparison.Ordinal);
            Assert.Contains("RibbonSyncController?.InvalidateRefreshState();", addInCodeText, StringComparison.Ordinal);
        }

        [Fact]
        public void ThisAddInRefreshesRibbonProjectWhenActiveSheetChanges()
        {
            var addInCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "ThisAddIn.cs"));

            Assert.Contains("Application.SheetActivate += Application_SheetActivate;", addInCodeText, StringComparison.Ordinal);
            Assert.Contains("Application.SheetActivate -= Application_SheetActivate;", addInCodeText, StringComparison.Ordinal);
            Assert.Contains("private void Application_SheetActivate(object sh)", addInCodeText, StringComparison.Ordinal);
            Assert.Contains("var sheetName = GetWorksheetName(sh);", addInCodeText, StringComparison.Ordinal);
            Assert.Contains("RibbonSyncController?.RefreshProjectFromSheetMetadata(sheetName);", addInCodeText, StringComparison.Ordinal);
            Assert.Contains("lastProjectRefreshSheetName = sheetName;", addInCodeText, StringComparison.Ordinal);
        }

        [Fact]
        public void ThisAddInRefreshesRibbonProjectWhenWorkbookActivatesAfterStartup()
        {
            var addInCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "ThisAddIn.cs"));

            Assert.Contains("Application.WorkbookActivate += Application_WorkbookActivate;", addInCodeText, StringComparison.Ordinal);
            Assert.Contains("Application.WorkbookActivate -= Application_WorkbookActivate;", addInCodeText, StringComparison.Ordinal);

            var methodStart = addInCodeText.IndexOf(
                "private void Application_WorkbookActivate(ExcelInterop.Workbook wb)",
                StringComparison.Ordinal);
            var nextMethodStart = addInCodeText.IndexOf(
                "private void Application_SheetChange(object sh, ExcelInterop.Range target)",
                methodStart,
                StringComparison.Ordinal);

            Assert.True(methodStart >= 0);
            Assert.True(nextMethodStart > methodStart);

            var methodBody = addInCodeText.Substring(methodStart, nextMethodStart - methodStart);
            Assert.Contains("RibbonSyncController?.InvalidateRefreshState();", methodBody, StringComparison.Ordinal);
            Assert.Contains("RibbonSyncController?.RefreshActiveProjectFromSheetMetadata();", methodBody, StringComparison.Ordinal);
            Assert.Contains("lastProjectRefreshSheetName = GetActiveWorksheetName();", methodBody, StringComparison.Ordinal);
        }

        [Fact]
        public void ThisAddInBindsRibbonToControllerAfterStartupInitialization()
        {
            var addInCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "ThisAddIn.cs"));

            Assert.Contains("Globals.Ribbons", addInCodeText, StringComparison.Ordinal);
            Assert.Contains("BindToControllersAndRefresh()", addInCodeText, StringComparison.Ordinal);
        }

        [Fact]
        public void ThisAddInIgnoresSelectionChangeEventsFromNonActiveSheets()
        {
            var addInCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "ThisAddIn.cs"));

            Assert.Contains("var activeSheetName = GetActiveWorksheetName();", addInCodeText, StringComparison.Ordinal);
            Assert.Contains("!string.Equals(sheetName, activeSheetName, StringComparison.OrdinalIgnoreCase)", addInCodeText, StringComparison.Ordinal);
            Assert.Contains("OfficeAgentLog.Info(\"excel\", \"selection.changed\", \"Excel selection changed.\");", addInCodeText, StringComparison.Ordinal);
        }

        [Fact]
        public void RibbonControllerDoesNotAutoInitializeWhenProjectIsSelected()
        {
            var ribbonControllerText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "RibbonSyncController.cs"));

            Assert.DoesNotContain("TryAutoInitializeCurrentSheet(sheetName, project);", ribbonControllerText, StringComparison.Ordinal);
        }

        [Fact]
        public void ProjectDropDownLabelsIncludeProjectIdPrefix()
        {
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            Assert.Contains("project?.ProjectId ?? string.Empty", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("project?.DisplayName ?? string.Empty", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("-", ribbonCodeText, StringComparison.Ordinal);
        }

        [Fact]
        public void RefreshProjectDropDownFormatsSelectedProjectWhenCurrentListDoesNotContainIt()
        {
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            Assert.Contains(
                "FormatProjectDropDownLabel(syncController.ActiveProjectId, syncController.ActiveProjectDisplayName)",
                ribbonCodeText,
                StringComparison.Ordinal);
        }

        [Fact]
        public void NoProjectRestoreTextUsesLastControllerOwnedStatusWhenNoItemsAndNoActiveProject()
        {
            var addInAssembly = Assembly.LoadFrom(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "bin",
                "Debug",
                "OfficeAgent.ExcelAddIn.dll"));
            var ribbonType = addInAssembly.GetType("OfficeAgent.ExcelAddIn.AgentRibbon", throwOnError: true);
            var method = ribbonType.GetMethod("GetNoProjectRestoreText", BindingFlags.Static | BindingFlags.NonPublic);

            Assert.NotNull(method);
            Assert.Equal(
                "请先登录",
                (string)method.Invoke(null, new object[] { 0, string.Empty, "请先登录" }));
            Assert.Equal(
                "先选择项目",
                (string)method.Invoke(null, new object[] { 0, string.Empty, string.Empty }));
            Assert.Equal(
                "先选择项目",
                (string)method.Invoke(null, new object[] { 0, string.Empty, "proj-a-项目A" }));
            Assert.Null(method.Invoke(null, new object[] { 1, string.Empty, "请先登录" }));
            Assert.Null(method.Invoke(null, new object[] { 0, "project-1", "请先登录" }));
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
