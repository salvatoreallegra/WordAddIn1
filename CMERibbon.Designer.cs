
namespace WordAddIn1
{
    partial class CMERibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public CMERibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CMERibbon));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.groupProcessComments = this.Factory.CreateRibbonGroup();
            this.toggleButton1 = this.Factory.CreateRibbonToggleButton();
            this.btnLoadComments = this.Factory.CreateRibbonButton();
            this.dDownComments = this.Factory.CreateRibbonDropDown();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.groupProcessComments.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.groupProcessComments);
            this.tab1.Label = "Compliance Made Easy";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.toggleButton1);
            this.group1.Label = "CME Task Pane Functions";
            this.group1.Name = "group1";
            // 
            // groupProcessComments
            // 
            this.groupProcessComments.Items.Add(this.dDownComments);
            this.groupProcessComments.Items.Add(this.btnLoadComments);
            this.groupProcessComments.Label = "CME Comment Functions";
            this.groupProcessComments.Name = "groupProcessComments";
            // 
            // toggleButton1
            // 
            this.toggleButton1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.toggleButton1.Image = ((System.Drawing.Image)(resources.GetObject("toggleButton1.Image")));
            this.toggleButton1.Label = "Toggle CME Task Pane";
            this.toggleButton1.Name = "toggleButton1";
            this.toggleButton1.ShowImage = true;
            this.toggleButton1.Tag = "";
            this.toggleButton1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButton1_Click);
            // 
            // btnLoadComments
            // 
            this.btnLoadComments.Label = "Load Comments";
            this.btnLoadComments.Name = "btnLoadComments";
            this.btnLoadComments.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoadComments_Click);
            // 
            // dDownComments
            // 
            this.dDownComments.Label = "All Comments";
            this.dDownComments.Name = "dDownComments";
            // 
            // CMERibbon
            // 
            this.Name = "CMERibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.groupProcessComments.ResumeLayout(false);
            this.groupProcessComments.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButton1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupProcessComments;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoadComments;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dDownComments;
    }

    partial class ThisRibbonCollection
    {
        internal CMERibbon Ribbon1
        {
            get { return this.GetRibbon<CMERibbon>(); }
        }
    }
}
