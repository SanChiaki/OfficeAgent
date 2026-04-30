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
        public void TaskPaneButtonUsesBuiltInOfficeImage()
        {
            var designerText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.Designer.cs"));
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            Assert.Contains("this.toggleTaskPaneButton.OfficeImageId = \"Info\";", designerText, StringComparison.Ordinal);
            Assert.Contains("this.toggleTaskPaneButton.ShowImage = true;", designerText, StringComparison.Ordinal);
            Assert.Contains("this.toggleTaskPaneButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;", designerText, StringComparison.Ordinal);
            Assert.Contains("this.toggleTaskPaneButton.Label = string.Empty;", designerText, StringComparison.Ordinal);
            Assert.Contains("this.toggleTaskPaneButton.ShowLabel = false;", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("this.toggleTaskPaneButton.Label = \"Open\";", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("this.toggleTaskPaneButton.ShowImage = false;", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("this.toggleTaskPaneButton.ShowLabel = true;", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("toggleTaskPaneButton.Image = Properties.Resources.Logo;", ribbonCodeText, StringComparison.Ordinal);
        }

        [Fact]
        public void RibbonButtonsUseSemanticBuiltInOfficeImages()
        {
            var designerText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.Designer.cs"));

            Assert.Contains("this.toggleTaskPaneButton.OfficeImageId = \"Info\";", designerText, StringComparison.Ordinal);
            Assert.Contains("this.initializeSheetButton.OfficeImageId = \"TableInsert\";", designerText, StringComparison.Ordinal);
            Assert.Contains("this.applyTemplateButton.OfficeImageId = \"FileOpen\";", designerText, StringComparison.Ordinal);
            Assert.Contains("this.saveTemplateButton.OfficeImageId = \"FileSave\";", designerText, StringComparison.Ordinal);
            Assert.Contains("this.saveAsTemplateButton.OfficeImageId = \"FileSaveAs\";", designerText, StringComparison.Ordinal);
            Assert.Contains("this.fullDownloadButton.OfficeImageId = \"RefreshAll\";", designerText, StringComparison.Ordinal);
            Assert.Contains("this.partialDownloadButton.OfficeImageId = \"Refresh\";", designerText, StringComparison.Ordinal);
            Assert.Contains("this.fullUploadButton.OfficeImageId = \"FilePublishToSharePoint\";", designerText, StringComparison.Ordinal);
            Assert.Contains("this.partialUploadButton.OfficeImageId = \"FileSendAsAttachment\";", designerText, StringComparison.Ordinal);
            Assert.Contains("this.loginButton.OfficeImageId = \"Lock\";", designerText, StringComparison.Ordinal);
            Assert.Contains("this.documentationButton.OfficeImageId = \"FileOpen\";", designerText, StringComparison.Ordinal);
            Assert.Contains("this.aboutButton.OfficeImageId = \"Info\";", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("ShowImage = false;", designerText, StringComparison.Ordinal);
        }

        [Fact]
        public void RibbonButtonsExplicitlyShowTheirOfficeImages()
        {
            var designerText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.Designer.cs"));

            Assert.Contains("this.toggleTaskPaneButton.ShowImage = true;", designerText, StringComparison.Ordinal);
            Assert.Contains("this.initializeSheetButton.ShowImage = true;", designerText, StringComparison.Ordinal);
            Assert.Contains("this.applyTemplateButton.ShowImage = true;", designerText, StringComparison.Ordinal);
            Assert.Contains("this.saveTemplateButton.ShowImage = true;", designerText, StringComparison.Ordinal);
            Assert.Contains("this.saveAsTemplateButton.ShowImage = true;", designerText, StringComparison.Ordinal);
            Assert.Contains("this.fullDownloadButton.ShowImage = true;", designerText, StringComparison.Ordinal);
            Assert.Contains("this.partialDownloadButton.ShowImage = true;", designerText, StringComparison.Ordinal);
            Assert.Contains("this.fullUploadButton.ShowImage = true;", designerText, StringComparison.Ordinal);
            Assert.Contains("this.partialUploadButton.ShowImage = true;", designerText, StringComparison.Ordinal);
            Assert.Contains("this.loginButton.ShowImage = true;", designerText, StringComparison.Ordinal);
            Assert.Contains("this.documentationButton.ShowImage = true;", designerText, StringComparison.Ordinal);
            Assert.Contains("this.aboutButton.ShowImage = true;", designerText, StringComparison.Ordinal);
        }

        [Fact]
        public void RibbonButtonsUseConfiguredLargeOrSmallLayouts()
        {
            var designerText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.Designer.cs"));

            Assert.Contains("this.toggleTaskPaneButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;", designerText, StringComparison.Ordinal);
            Assert.Contains("this.initializeSheetButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular;", designerText, StringComparison.Ordinal);
            Assert.Contains("this.applyTemplateButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular;", designerText, StringComparison.Ordinal);
            Assert.Contains("this.saveTemplateButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular;", designerText, StringComparison.Ordinal);
            Assert.Contains("this.saveAsTemplateButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular;", designerText, StringComparison.Ordinal);
            Assert.Contains("this.fullDownloadButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;", designerText, StringComparison.Ordinal);
            Assert.Contains("this.partialDownloadButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;", designerText, StringComparison.Ordinal);
            Assert.Contains("this.fullUploadButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;", designerText, StringComparison.Ordinal);
            Assert.Contains("this.partialUploadButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;", designerText, StringComparison.Ordinal);
            Assert.Contains("this.loginButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;", designerText, StringComparison.Ordinal);
            Assert.Contains("this.documentationButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;", designerText, StringComparison.Ordinal);
            Assert.Contains("this.aboutButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;", designerText, StringComparison.Ordinal);
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
        public void RibbonTabAndAgentGroupUseXisdpBrandingWhileAgentButtonUsesIconOnly()
        {
            var designerText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.Designer.cs"));
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            Assert.Contains("this.tab1.Label = \"xISDP\";", designerText, StringComparison.Ordinal);
            Assert.Contains("this.group1.Label = \"xISDP AI\";", designerText, StringComparison.Ordinal);
            Assert.Contains("this.toggleTaskPaneButton.Label = string.Empty;", designerText, StringComparison.Ordinal);
            Assert.Contains("this.toggleTaskPaneButton.ShowLabel = false;", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("this.toggleTaskPaneButton.Label = \"Open\";", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("toggleTaskPaneButton.Label = \"Open\";", ribbonCodeText, StringComparison.Ordinal);
            Assert.DoesNotContain("this.tab1.Label = \"X-ISDP\";", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("this.tab1.Label = \"ISDP\";", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("this.group1.Label = \"ISDP AI\";", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("this.toggleTaskPaneButton.Label = \"ISDP AI\";", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("Resy AI", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("Resy AI", ribbonCodeText, StringComparison.Ordinal);
        }

        [Fact]
        public void TaskPaneTitleUsesXisdpAiBranding()
        {
            var taskPaneControllerText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "TaskPane",
                "TaskPaneController.cs"));

            Assert.Contains("addIn.CustomTaskPanes.Add(hostControl, \"xISDP AI\")", taskPaneControllerText, StringComparison.Ordinal);
            Assert.DoesNotContain("addIn.CustomTaskPanes.Add(hostControl, \"ISDP AI\")", taskPaneControllerText, StringComparison.Ordinal);
        }

        [Fact]
        public void RibbonDesignerUsesEnglishSafeDefaultsForLocalizedControls()
        {
            var designerText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.Designer.cs"));

            Assert.Contains("this.groupProject.Label = \"Project\";", designerText, StringComparison.Ordinal);
            Assert.Contains("this.projectSelectorButton.Label = \"Select project\";", designerText, StringComparison.Ordinal);
            Assert.Contains("this.initializeSheetButton.Label = \"Initialize sheet\";", designerText, StringComparison.Ordinal);
            Assert.Contains("this.groupTemplate.Label = \"Setting\";", designerText, StringComparison.Ordinal);
            Assert.Contains("this.applyTemplateButton.Label = \"Apply Setting\";", designerText, StringComparison.Ordinal);
            Assert.Contains("this.saveTemplateButton.Label = \"Save Setting\";", designerText, StringComparison.Ordinal);
            Assert.Contains("this.saveAsTemplateButton.Label = \"Save as Setting\";", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("this.applyTemplateButton.Label = \"Apply setting\";", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("this.saveTemplateButton.Label = \"Save setting\";", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("this.saveAsTemplateButton.Label = \"Save as setting\";", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("this.groupTemplate.Label = \"Config\";", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("this.applyTemplateButton.Label = \"Apply config\";", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("this.saveTemplateButton.Label = \"Save config\";", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("this.saveAsTemplateButton.Label = \"Save as config\";", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("this.groupTemplate.Label = \"Template\";", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("this.applyTemplateButton.Label = \"Apply template\";", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("this.saveTemplateButton.Label = \"Save template\";", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("this.saveAsTemplateButton.Label = \"Save as template\";", designerText, StringComparison.Ordinal);
            Assert.Contains("this.groupDataSync.Label = \"Data sync\";", designerText, StringComparison.Ordinal);
            Assert.Contains("this.partialDownloadButton.Label = \"Download\";", designerText, StringComparison.Ordinal);
            Assert.Contains("this.partialUploadButton.Label = \"Upload\";", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("this.partialDownloadButton.Label = \"Partial download\";", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("this.partialUploadButton.Label = \"Partial upload\";", designerText, StringComparison.Ordinal);
            Assert.Contains("this.group2.Label = \"Account\";", designerText, StringComparison.Ordinal);
            Assert.Contains("this.loginButton.Label = \"Login\";", designerText, StringComparison.Ordinal);
            Assert.Contains("this.groupHelp.Label = \"Help\";", designerText, StringComparison.Ordinal);
            Assert.Contains("this.documentationButton.Label = \"Documentation\";", designerText, StringComparison.Ordinal);
            Assert.Contains("this.aboutButton.Label = \"About\";", designerText, StringComparison.Ordinal);
        }

        [Fact]
        public void DataSyncGroupContainsPartialDownloadAndUploadOnly()
        {
            var designerText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.Designer.cs"));

            Assert.Contains("this.groupDataSync.Label = \"Data sync\";", designerText, StringComparison.Ordinal);
            Assert.Contains("this.groupDataSync.Items.Add(this.partialDownloadButton);", designerText, StringComparison.Ordinal);
            Assert.Contains("this.groupDataSync.Items.Add(this.partialUploadButton);", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("this.groupDataSync.Items.Add(this.fullDownloadButton);", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("this.groupDataSync.Items.Add(this.fullUploadButton);", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("this.groupDownload", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("this.groupUpload", designerText, StringComparison.Ordinal);
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
        public void TemplateGroupAppearsAfterProjectGroupAndBeforeDataSyncGroup()
        {
            var designerText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.Designer.cs"));

            var projectGroupIndex = designerText.IndexOf("this.tab1.Groups.Add(this.groupProject);", StringComparison.Ordinal);
            var templateGroupIndex = designerText.IndexOf("this.tab1.Groups.Add(this.groupTemplate);", StringComparison.Ordinal);
            var dataSyncGroupIndex = designerText.IndexOf("this.tab1.Groups.Add(this.groupDataSync);", StringComparison.Ordinal);

            Assert.True(projectGroupIndex >= 0);
            Assert.True(templateGroupIndex > projectGroupIndex);
            Assert.True(dataSyncGroupIndex > templateGroupIndex);
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
            Assert.DoesNotContain("templateActionsBox", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("templateSaveButtonsBox", designerText, StringComparison.Ordinal);
        }

        [Fact]
        public void HelpGroupContainsDocumentationAndAboutButtons()
        {
            var designerText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.Designer.cs"));

            Assert.Contains("this.tab1.Groups.Add(this.groupHelp);", designerText, StringComparison.Ordinal);
            Assert.Contains("this.groupHelp.Label = \"Help\";", designerText, StringComparison.Ordinal);
            Assert.Contains("this.groupHelp.Items.Add(this.documentationButton);", designerText, StringComparison.Ordinal);
            Assert.Contains("this.groupHelp.Items.Add(this.aboutButton);", designerText, StringComparison.Ordinal);
            Assert.Contains("this.documentationButton.Label = \"Documentation\";", designerText, StringComparison.Ordinal);
            Assert.Contains("this.aboutButton.Label = \"About\";", designerText, StringComparison.Ordinal);
            Assert.Contains("this.documentationButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DocumentationButton_Click);", designerText, StringComparison.Ordinal);
            Assert.Contains("this.aboutButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AboutButton_Click);", designerText, StringComparison.Ordinal);
        }

        [Fact]
        public void DocumentationButtonOpensConfiguredDocumentationUrlInDefaultBrowser()
        {
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            Assert.Contains("private const string DocumentationUrl = \"https://github.com/SanChiaki/OfficeAgent\";", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("private static void OpenUrlInDefaultBrowser(string url)", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("ProcessStartInfo", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("UseShellExecute = true", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("OpenUrlInDefaultBrowser(DocumentationUrl);", ribbonCodeText, StringComparison.Ordinal);
        }

        [Fact]
        public void AboutButtonShowsVersionAndBuildInformation()
        {
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            Assert.Contains("private static string CreateAboutMessage()", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("VersionInfo.AppVersion", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("GetBuildConfiguration()", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("File.GetLastWriteTime", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("MessageBox.Show(CreateAboutMessage()", ribbonCodeText, StringComparison.Ordinal);
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
            Assert.Contains("ShowAuthenticationRequired(GetStrings().AuthenticationRequiredDefaultMessage)", ribbonCodeText, StringComparison.Ordinal);
            Assert.DoesNotContain("ShowAuthenticationRequired(ex.Message)", ribbonCodeText, StringComparison.Ordinal);
        }

        [Fact]
        public void ProjectLoadingMarksDropdownAsLoginRequiredWhenAuthenticationFails()
        {
            var addInAssembly = Assembly.LoadFrom(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "bin",
                "Debug",
                "OfficeAgent.ExcelAddIn.dll"));
            var ribbonType = addInAssembly.GetType("OfficeAgent.ExcelAddIn.AgentRibbon", throwOnError: true);
            var method = ribbonType.GetMethod(
                "GetNoProjectRestoreText",
                BindingFlags.Static | BindingFlags.NonPublic,
                binder: null,
                types: new[] { typeof(int), typeof(string), typeof(string) },
                modifiers: null);

            Assert.NotNull(method);
            Assert.Equal("请先登录", (string)method.Invoke(null, new object[] { 0, string.Empty, "请先登录" }));
            Assert.Equal("Sign in first", (string)method.Invoke(null, new object[] { 0, string.Empty, "Sign in first" }));
        }

        [Fact]
        public void AuthenticationPromptOffersPointMeToLoginButton()
        {
            var dialogCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "Dialogs",
                "OperationResultDialog.cs"));

            Assert.Contains("AuthenticationRequiredLoginButtonText", dialogCodeText, StringComparison.Ordinal);
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
            var addInAssembly = Assembly.LoadFrom(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "bin",
                "Debug",
                "OfficeAgent.ExcelAddIn.dll"));
            var stringsType = addInAssembly.GetType("OfficeAgent.ExcelAddIn.Localization.HostLocalizedStrings", throwOnError: true);
            var forLocale = stringsType.GetMethod("ForLocale", BindingFlags.Public | BindingFlags.Static);
            var zhStrings = forLocale.Invoke(null, new object[] { "zh" });
            var enStrings = forLocale.Invoke(null, new object[] { "en" });
            var messageMethod = stringsType.GetMethod("ProjectListLoadFailedMessage", BindingFlags.Instance | BindingFlags.Public);

            Assert.NotNull(messageMethod);
            Assert.Contains("项目列表加载失败", (string)messageMethod.Invoke(zhStrings, new object[] { "boom" }), StringComparison.Ordinal);
            Assert.Contains("Failed to load projects", (string)messageMethod.Invoke(enStrings, new object[] { "boom" }), StringComparison.Ordinal);
        }

        [Fact]
        public void ProjectLoadingUsesDedicatedStatusItemsInsteadOfRibbonLabelOnly()
        {
            var addInAssembly = Assembly.LoadFrom(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "bin",
                "Debug",
                "OfficeAgent.ExcelAddIn.dll"));
            var stringsType = addInAssembly.GetType("OfficeAgent.ExcelAddIn.Localization.HostLocalizedStrings", throwOnError: true);
            var forLocale = stringsType.GetMethod("ForLocale", BindingFlags.Public | BindingFlags.Static);
            var zhStrings = forLocale.Invoke(null, new object[] { "zh" });
            var enStrings = forLocale.Invoke(null, new object[] { "en" });

            Assert.Equal("请先登录", stringsType.GetProperty("ProjectDropDownLoginRequiredText").GetValue(zhStrings));
            Assert.Equal("无可用项目", stringsType.GetProperty("ProjectDropDownNoAvailableProjectsText").GetValue(zhStrings));
            Assert.Equal("Sign in first", stringsType.GetProperty("ProjectDropDownLoginRequiredText").GetValue(enStrings));
            Assert.Equal("No projects available", stringsType.GetProperty("ProjectDropDownNoAvailableProjectsText").GetValue(enStrings));
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
            var addInAssembly = Assembly.LoadFrom(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "bin",
                "Debug",
                "OfficeAgent.ExcelAddIn.dll"));
            var stringsType = addInAssembly.GetType("OfficeAgent.ExcelAddIn.Localization.HostLocalizedStrings", throwOnError: true);
            var forLocale = stringsType.GetMethod("ForLocale", BindingFlags.Public | BindingFlags.Static);
            var zhStrings = forLocale.Invoke(null, new object[] { "zh" });
            var enStrings = forLocale.Invoke(null, new object[] { "en" });

            Assert.Equal("先选择项目", stringsType.GetProperty("ProjectDropDownPlaceholderText").GetValue(zhStrings));
            Assert.Equal("Select project", stringsType.GetProperty("ProjectDropDownPlaceholderText").GetValue(enStrings));
        }

        [Fact]
        public void PopulateProjectDropDownCachesLoadedProjectsForCustomPicker()
        {
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            Assert.Contains("projectSelectorEntries.Clear();", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("projectSelectorEntries.Add(new ProjectSelectorEntry", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("Label = projectLabel,", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("Tag = projectKey,", ribbonCodeText, StringComparison.Ordinal);
        }

        [Fact]
        public void ProjectSelectorDisplaysCurrentTextOnRibbonButton()
        {
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            Assert.Contains("projectSelectorButton.Label = normalizedText;", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("projectSelectorButton.ScreenTip = normalizedText;", ribbonCodeText, StringComparison.Ordinal);
            Assert.DoesNotContain("projectDropDown.SelectedItem", ribbonCodeText, StringComparison.Ordinal);
        }

        [Fact]
        public void ProjectSelectorInvalidatesRibbonControlAfterProgrammaticSelectionChanges()
        {
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            Assert.Contains("RibbonUI?.InvalidateControl(projectSelectorButton.Name);", ribbonCodeText, StringComparison.Ordinal);
        }

        [Fact]
        public void ProjectSelectorUsesButtonClickToOpenCustomPicker()
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
                "this.projectSelectorButton = Factory.CreateRibbonButton();",
                designerText,
                StringComparison.Ordinal);
            Assert.Contains(
                "this.groupProject.Items.Add(this.projectSelectorButton);",
                designerText,
                StringComparison.Ordinal);
            Assert.Contains(
                "this.projectSelectorButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ProjectSelectorButton_Click);",
                designerText,
                StringComparison.Ordinal);
            Assert.DoesNotContain("Factory.CreateRibbonDropDown()", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("this.projectDropDown.ItemsLoading +=", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("this.projectDropDown.SelectionChanged +=", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("this.projectDropDown.ButtonClick +=", designerText, StringComparison.Ordinal);
            Assert.DoesNotContain("projectSearchBox", designerText, StringComparison.Ordinal);
            Assert.Contains("private void ProjectSelectorButton_Click(object sender, RibbonControlEventArgs e)", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("ShowProjectPickerDialog();", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("PopulateProjectDropDown();", ribbonCodeText, StringComparison.Ordinal);
            Assert.Contains("RefreshProjectDropDownFromController();", ribbonCodeText, StringComparison.Ordinal);
        }

        [Fact]
        public void ProjectPickerDialogUsesRealtimeFuzzySearch()
        {
            var dialogText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "Dialogs",
                "ProjectPickerDialog.cs"));

            Assert.Contains("private readonly TextBox searchTextBox;", dialogText, StringComparison.Ordinal);
            Assert.Contains("private readonly ListBox projectListBox;", dialogText, StringComparison.Ordinal);
            Assert.Contains("searchTextBox.TextChanged += SearchTextBox_TextChanged;", dialogText, StringComparison.Ordinal);
            Assert.Contains("ProjectSearchMatcher.IsMatch(item.Label, searchTextBox.Text)", dialogText, StringComparison.Ordinal);
            Assert.Contains("public ProjectOption SelectedProject", dialogText, StringComparison.Ordinal);
        }

        [Fact]
        public void ActiveProjectChangeRefreshesProjectSelectorText()
        {
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            var methodStart = ribbonCodeText.IndexOf("private void SyncController_ActiveProjectChanged(object sender, EventArgs e)", StringComparison.Ordinal);
            var nextMethodStart = ribbonCodeText.IndexOf("private void TemplateController_TemplateStateChanged(object sender, EventArgs e)", methodStart, StringComparison.Ordinal);

            Assert.True(methodStart >= 0);
            Assert.True(nextMethodStart > methodStart);

            var methodBody = ribbonCodeText.Substring(methodStart, nextMethodStart - methodStart);
            Assert.Contains("RefreshProjectDropDownFromController();", methodBody, StringComparison.Ordinal);
            Assert.DoesNotContain("RebuildProjectDropDownItemsFromCurrentState();", methodBody, StringComparison.Ordinal);
            Assert.DoesNotContain("ResetProjectDropDownItemsToPlaceholderOnly();", methodBody, StringComparison.Ordinal);
        }

        [Fact]
        public void ProjectPickerSelectionGoesThroughControllerSelectProject()
        {
            var ribbonCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "AgentRibbon.cs"));

            var methodStart = ribbonCodeText.IndexOf("private void ShowProjectPickerDialog()", StringComparison.Ordinal);
            var nextMethodStart = ribbonCodeText.IndexOf("internal void BindToSyncControllerAndRefresh()", methodStart, StringComparison.Ordinal);

            Assert.True(methodStart >= 0);
            Assert.True(nextMethodStart > methodStart);

            var methodBody = ribbonCodeText.Substring(methodStart, nextMethodStart - methodStart);
            Assert.Contains("dialog.SelectedProject", methodBody, StringComparison.Ordinal);
            Assert.Contains("Globals.ThisAddIn.RibbonSyncController?.SelectProject(dialog.SelectedProject);", methodBody, StringComparison.Ordinal);
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
            var metadataNamesText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "Excel",
                "MetadataWorksheetNames.cs"));

            Assert.Contains("Application.SheetChange += Application_SheetChange;", addInCodeText, StringComparison.Ordinal);
            Assert.Contains("private void Application_SheetChange(object sh, ExcelInterop.Range target)", addInCodeText, StringComparison.Ordinal);
            Assert.Contains("IsSettingsSheet(sheetName)", addInCodeText, StringComparison.Ordinal);
            Assert.Contains("MetadataWorksheetNames.IsMetadataWorksheet(sheetName)", addInCodeText, StringComparison.Ordinal);
            Assert.Contains("xISDP_Setting", metadataNamesText, StringComparison.Ordinal);
            Assert.Contains("ISDP_Setting", metadataNamesText, StringComparison.Ordinal);
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
            var method = ribbonType.GetMethod(
                "GetNoProjectRestoreText",
                BindingFlags.Static | BindingFlags.NonPublic,
                binder: null,
                types: new[] { typeof(int), typeof(string), typeof(string) },
                modifiers: null);

            Assert.NotNull(method);
            Assert.Equal(
                "Sign in first",
                (string)method.Invoke(null, new object[] { 0, string.Empty, "Sign in first" }));
            Assert.Equal(
                "无可用项目",
                (string)method.Invoke(null, new object[] { 0, string.Empty, "无可用项目" }));
            Assert.Equal(
                "Select project",
                (string)method.Invoke(null, new object[] { 0, string.Empty, string.Empty }));
            Assert.Equal(
                "Select project",
                (string)method.Invoke(null, new object[] { 0, string.Empty, "proj-a-项目A" }));
            Assert.Null(method.Invoke(null, new object[] { 1, string.Empty, "Sign in first" }));
            Assert.Null(method.Invoke(null, new object[] { 0, "project-1", "Sign in first" }));
        }

        [Fact]
        public void ThisAddInTracksPendingCellEditsForWorkbookChangeLog()
        {
            var addInCodeText = File.ReadAllText(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "ThisAddIn.cs"));

            Assert.Contains("internal WorksheetPendingEditTracker WorksheetPendingEditTracker { get; private set; }", addInCodeText, StringComparison.Ordinal);
            Assert.Contains("internal IWorksheetChangeLogStore WorksheetChangeLogStore { get; private set; }", addInCodeText, StringComparison.Ordinal);
            Assert.Contains("WorksheetChangeLogStore = new WorksheetChangeLogStore(worksheetGridAdapter);", addInCodeText, StringComparison.Ordinal);
            Assert.Contains("WorksheetPendingEditTracker = new WorksheetPendingEditTracker();", addInCodeText, StringComparison.Ordinal);
            Assert.Contains("WorksheetPendingEditTracker.CaptureBeforeValues(sheetName, ReadWorksheetCellValues(target));", addInCodeText, StringComparison.Ordinal);
            Assert.Contains("WorksheetPendingEditTracker.MarkChanged(sheetName, ReadWorksheetCellAddresses(target));", addInCodeText, StringComparison.Ordinal);
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
