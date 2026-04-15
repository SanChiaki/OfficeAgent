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
            this.groupDownload = Factory.CreateRibbonGroup();
            this.groupUpload = Factory.CreateRibbonGroup();
            this.group2 = Factory.CreateRibbonGroup();
            this.toggleTaskPaneButton = Factory.CreateRibbonButton();
            this.projectDropDown = Factory.CreateRibbonDropDown();
            this.initializeSheetButton = Factory.CreateRibbonButton();
            this.fullDownloadButton = Factory.CreateRibbonButton();
            this.partialDownloadButton = Factory.CreateRibbonButton();
            this.fullUploadButton = Factory.CreateRibbonButton();
            this.partialUploadButton = Factory.CreateRibbonButton();
            this.loginButton = Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.groupProject.SuspendLayout();
            this.groupDownload.SuspendLayout();
            this.groupUpload.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.groupProject);
            this.tab1.Groups.Add(this.groupDownload);
            this.tab1.Groups.Add(this.groupUpload);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "Resy AI";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.toggleTaskPaneButton);
            this.group1.Label = "Resy AI";
            this.group1.Name = "groupAgent";
            // 
            // toggleTaskPaneButton
            // 
            this.toggleTaskPaneButton.Label = "Resy AI";
            this.toggleTaskPaneButton.Name = "openTaskPaneButton";
            this.toggleTaskPaneButton.ShowImage = false;
            this.toggleTaskPaneButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ToggleTaskPaneButton_Click);
            // 
            // groupProject
            // 
            this.groupProject.Items.Add(this.projectDropDown);
            this.groupProject.Items.Add(this.initializeSheetButton);
            this.groupProject.Label = "\u9879\u76EE";
            this.groupProject.Name = "groupProject";
            // 
            // projectDropDown
            // 
            this.projectDropDown.Label = "\u5148\u9009\u62E9\u9879\u76EE";
            this.projectDropDown.Name = "projectDropDown";
            this.projectDropDown.ShowItemLabel = true;
            this.projectDropDown.ShowLabel = false;
            this.projectDropDown.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ProjectDropDown_SelectionChanged);
            // 
            // initializeSheetButton
            // 
            this.initializeSheetButton.Label = "\u521D\u59CB\u5316\u5F53\u524D\u8868";
            this.initializeSheetButton.Name = "initializeSheetButton";
            this.initializeSheetButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.InitializeSheetButton_Click);
            // 
            // groupDownload
            // 
            this.groupDownload.Items.Add(this.fullDownloadButton);
            this.groupDownload.Items.Add(this.partialDownloadButton);
            this.groupDownload.Label = "\u4E0B\u8F7D";
            this.groupDownload.Name = "groupDownload";
            // 
            // fullDownloadButton
            // 
            this.fullDownloadButton.Label = "\u5168\u91CF\u4E0B\u8F7D";
            this.fullDownloadButton.Name = "fullDownloadButton";
            this.fullDownloadButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.FullDownloadButton_Click);
            // 
            // partialDownloadButton
            // 
            this.partialDownloadButton.Label = "\u90E8\u5206\u4E0B\u8F7D";
            this.partialDownloadButton.Name = "partialDownloadButton";
            this.partialDownloadButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.PartialDownloadButton_Click);
            // 
            // groupUpload
            // 
            this.groupUpload.Items.Add(this.fullUploadButton);
            this.groupUpload.Items.Add(this.partialUploadButton);
            this.groupUpload.Label = "\u4E0A\u4F20";
            this.groupUpload.Name = "groupUpload";
            // 
            // fullUploadButton
            // 
            this.fullUploadButton.Label = "\u5168\u91CF\u4E0A\u4F20";
            this.fullUploadButton.Name = "fullUploadButton";
            this.fullUploadButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.FullUploadButton_Click);
            // 
            // partialUploadButton
            // 
            this.partialUploadButton.Label = "\u90E8\u5206\u4E0A\u4F20";
            this.partialUploadButton.Name = "partialUploadButton";
            this.partialUploadButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.PartialUploadButton_Click);
            //
            // loginButton
            //
            this.group2.Items.Add(this.loginButton);
            this.group2.Label = "\u8D26\u53F7";
            this.group2.Name = "group2";
            this.loginButton.Label = "\u767B\u5F55";
            this.loginButton.Name = "loginButton";
            this.loginButton.ShowImage = false;
            this.loginButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.LoginButton_Click);
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
            this.groupDownload.ResumeLayout(false);
            this.groupDownload.PerformLayout();
            this.groupUpload.ResumeLayout(false);
            this.groupUpload.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);
        }

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupProject;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupDownload;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupUpload;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton toggleTaskPaneButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown projectDropDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton initializeSheetButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton fullDownloadButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton partialDownloadButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton fullUploadButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton partialUploadButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton loginButton;
    }

    partial class ThisRibbonCollection
    {
        internal AgentRibbon AgentRibbon => this.GetRibbon<AgentRibbon>();
    }
}
