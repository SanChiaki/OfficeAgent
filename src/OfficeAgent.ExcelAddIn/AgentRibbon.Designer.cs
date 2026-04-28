namespace OfficeAgent.ExcelAddIn
{
    partial class AgentRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        private System.ComponentModel.IContainer components = null;

        public AgentRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }

            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.tab1 = Factory.CreateRibbonTab();
            this.group1 = Factory.CreateRibbonGroup();
            this.groupProject = Factory.CreateRibbonGroup();
            this.groupTemplate = Factory.CreateRibbonGroup();
            this.groupDataSync = Factory.CreateRibbonGroup();
            this.group2 = Factory.CreateRibbonGroup();
            this.groupHelp = Factory.CreateRibbonGroup();
            this.toggleTaskPaneButton = Factory.CreateRibbonButton();
            this.projectDropDown = Factory.CreateRibbonDropDown();
            this.initializeSheetButton = Factory.CreateRibbonButton();
            this.applyTemplateButton = Factory.CreateRibbonButton();
            this.saveTemplateButton = Factory.CreateRibbonButton();
            this.saveAsTemplateButton = Factory.CreateRibbonButton();
            this.fullDownloadButton = Factory.CreateRibbonButton();
            this.partialDownloadButton = Factory.CreateRibbonButton();
            this.fullUploadButton = Factory.CreateRibbonButton();
            this.partialUploadButton = Factory.CreateRibbonButton();
            this.loginButton = Factory.CreateRibbonButton();
            this.documentationButton = Factory.CreateRibbonButton();
            this.aboutButton = Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.groupProject.SuspendLayout();
            this.groupTemplate.SuspendLayout();
            this.groupDataSync.SuspendLayout();
            this.group2.SuspendLayout();
            this.groupHelp.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.groupProject);
            this.tab1.Groups.Add(this.groupTemplate);
            this.tab1.Groups.Add(this.groupDataSync);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.groupHelp);
            this.tab1.Label = "ISDP";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.toggleTaskPaneButton);
            this.group1.Label = "ISDP AI";
            this.group1.Name = "groupAgent";
            // 
            // toggleTaskPaneButton
            // 
            this.toggleTaskPaneButton.Label = "Open";
            this.toggleTaskPaneButton.Name = "openTaskPaneButton";
            this.toggleTaskPaneButton.OfficeImageId = "Info";
            this.toggleTaskPaneButton.ShowImage = true;
            this.toggleTaskPaneButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.toggleTaskPaneButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ToggleTaskPaneButton_Click);
            // 
            // groupProject
            // 
            this.groupProject.Items.Add(this.projectDropDown);
            this.groupProject.Items.Add(this.initializeSheetButton);
            this.groupProject.Label = "Project";
            this.groupProject.Name = "groupProject";
            // 
            // projectDropDown
            // 
            this.projectDropDown.Label = "Select project";
            this.projectDropDown.Name = "projectDropDown";
            this.projectDropDown.SizeString = "WWWWWWWWWWWWWWWWWWWW";
            this.projectDropDown.ShowItemLabel = true;
            this.projectDropDown.ShowLabel = false;
            this.projectDropDown.ItemsLoading += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ProjectDropDown_ItemsLoading);
            this.projectDropDown.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ProjectDropDown_SelectionChanged);
            // 
            // initializeSheetButton
            // 
            this.initializeSheetButton.Label = "Initialize sheet";
            this.initializeSheetButton.Name = "initializeSheetButton";
            this.initializeSheetButton.OfficeImageId = "TableInsert";
            this.initializeSheetButton.ShowImage = true;
            this.initializeSheetButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.initializeSheetButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.InitializeSheetButton_Click);
            // 
            // groupTemplate
            // 
            this.groupTemplate.Items.Add(this.applyTemplateButton);
            this.groupTemplate.Items.Add(this.saveTemplateButton);
            this.groupTemplate.Items.Add(this.saveAsTemplateButton);
            this.groupTemplate.Label = "Template";
            this.groupTemplate.Name = "groupTemplate";
            // 
            // applyTemplateButton
            // 
            this.applyTemplateButton.Label = "Apply template";
            this.applyTemplateButton.Name = "applyTemplateButton";
            this.applyTemplateButton.OfficeImageId = "FileOpen";
            this.applyTemplateButton.ShowImage = true;
            this.applyTemplateButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.applyTemplateButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ApplyTemplateButton_Click);
            // 
            // saveTemplateButton
            // 
            this.saveTemplateButton.Label = "Save template";
            this.saveTemplateButton.Name = "saveTemplateButton";
            this.saveTemplateButton.OfficeImageId = "FileSave";
            this.saveTemplateButton.ShowImage = true;
            this.saveTemplateButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.saveTemplateButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SaveTemplateButton_Click);
            // 
            // saveAsTemplateButton
            // 
            this.saveAsTemplateButton.Label = "Save as template";
            this.saveAsTemplateButton.Name = "saveAsTemplateButton";
            this.saveAsTemplateButton.OfficeImageId = "FileSaveAs";
            this.saveAsTemplateButton.ShowImage = true;
            this.saveAsTemplateButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.saveAsTemplateButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SaveAsTemplateButton_Click);
            // 
            // groupDataSync
            // 
            this.groupDataSync.Items.Add(this.partialDownloadButton);
            this.groupDataSync.Items.Add(this.partialUploadButton);
            this.groupDataSync.Label = "Data sync";
            this.groupDataSync.Name = "groupDataSync";
            // 
            // fullDownloadButton
            // 
            this.fullDownloadButton.Label = "Full download";
            this.fullDownloadButton.Name = "fullDownloadButton";
            this.fullDownloadButton.OfficeImageId = "RefreshAll";
            this.fullDownloadButton.ShowImage = true;
            this.fullDownloadButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.fullDownloadButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.FullDownloadButton_Click);
            // 
            // partialDownloadButton
            // 
            this.partialDownloadButton.Label = "Partial download";
            this.partialDownloadButton.Name = "partialDownloadButton";
            this.partialDownloadButton.OfficeImageId = "Refresh";
            this.partialDownloadButton.ShowImage = true;
            this.partialDownloadButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.partialDownloadButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.PartialDownloadButton_Click);
            // 
            // fullUploadButton
            // 
            this.fullUploadButton.Label = "Full upload";
            this.fullUploadButton.Name = "fullUploadButton";
            this.fullUploadButton.OfficeImageId = "FilePublishToSharePoint";
            this.fullUploadButton.ShowImage = true;
            this.fullUploadButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.fullUploadButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.FullUploadButton_Click);
            // 
            // partialUploadButton
            // 
            this.partialUploadButton.Label = "Partial upload";
            this.partialUploadButton.Name = "partialUploadButton";
            this.partialUploadButton.OfficeImageId = "FileSendAsAttachment";
            this.partialUploadButton.ShowImage = true;
            this.partialUploadButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.partialUploadButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.PartialUploadButton_Click);
            //
            // loginButton
            //
            this.group2.Items.Add(this.loginButton);
            this.group2.Label = "Account";
            this.group2.Name = "group2";
            this.loginButton.Label = "Login";
            this.loginButton.Name = "loginButton";
            this.loginButton.OfficeImageId = "Lock";
            this.loginButton.ShowImage = true;
            this.loginButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.loginButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.LoginButton_Click);
            //
            // groupHelp
            //
            this.groupHelp.Items.Add(this.documentationButton);
            this.groupHelp.Items.Add(this.aboutButton);
            this.groupHelp.Label = "Help";
            this.groupHelp.Name = "groupHelp";
            //
            // documentationButton
            //
            this.documentationButton.Label = "Documentation";
            this.documentationButton.Name = "documentationButton";
            this.documentationButton.OfficeImageId = "FileOpen";
            this.documentationButton.ShowImage = true;
            this.documentationButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.documentationButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DocumentationButton_Click);
            //
            // aboutButton
            //
            this.aboutButton.Label = "About";
            this.aboutButton.Name = "aboutButton";
            this.aboutButton.OfficeImageId = "Info";
            this.aboutButton.ShowImage = true;
            this.aboutButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.aboutButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AboutButton_Click);
            // 
            // AgentRibbon
            // 
            this.Name = "AgentRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.AgentRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.groupProject.ResumeLayout(false);
            this.groupProject.PerformLayout();
            this.groupTemplate.ResumeLayout(false);
            this.groupTemplate.PerformLayout();
            this.groupDataSync.ResumeLayout(false);
            this.groupDataSync.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.groupHelp.ResumeLayout(false);
            this.groupHelp.PerformLayout();
            this.ResumeLayout(false);
        }

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupProject;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupTemplate;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupDataSync;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupHelp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton toggleTaskPaneButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown projectDropDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton initializeSheetButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton applyTemplateButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton saveTemplateButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton saveAsTemplateButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton fullDownloadButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton partialDownloadButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton fullUploadButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton partialUploadButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton loginButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton documentationButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton aboutButton;
    }

    partial class ThisRibbonCollection
    {
        internal AgentRibbon AgentRibbon => this.GetRibbon<AgentRibbon>();
    }
}
