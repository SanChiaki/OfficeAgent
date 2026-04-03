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
            this.group2 = Factory.CreateRibbonGroup();
            this.toggleTaskPaneButton = Factory.CreateRibbonButton();
            this.loginButton = Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabAddIns";
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "Resy AI";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.toggleTaskPaneButton);
            this.group1.Label = "Resy AI";
            this.group1.Name = "group1";
            // 
            // toggleTaskPaneButton
            // 
            this.toggleTaskPaneButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.toggleTaskPaneButton.Label = "Open";
            this.toggleTaskPaneButton.Name = "toggleTaskPaneButton";
            this.toggleTaskPaneButton.ShowImage = true;
            this.toggleTaskPaneButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ToggleTaskPaneButton_Click);
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
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);
        }

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton toggleTaskPaneButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton loginButton;
    }

    partial class ThisRibbonCollection
    {
        internal AgentRibbon AgentRibbon => this.GetRibbon<AgentRibbon>();
    }
}
