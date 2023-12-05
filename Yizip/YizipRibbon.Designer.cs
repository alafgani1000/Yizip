namespace Yizip
{
    partial class YizipRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public YizipRibbon()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(YizipRibbon));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.YizipGroup = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.btnCache = this.Factory.CreateRibbonButton();
            this.ribbonZipStatus = this.Factory.CreateRibbonGroup();
            this.btnActive = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.YizipGroup.SuspendLayout();
            this.ribbonZipStatus.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.YizipGroup);
            this.tab1.Groups.Add(this.ribbonZipStatus);
            this.tab1.Label = "Yasunaga";
            this.tab1.Name = "tab1";
            // 
            // YizipGroup
            // 
            this.YizipGroup.Items.Add(this.button1);
            this.YizipGroup.Items.Add(this.btnCache);
            this.YizipGroup.Label = "Compress";
            this.YizipGroup.Name = "YizipGroup";
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.Label = "Yasunaga Zip";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button1_Click_1);
            // 
            // btnCache
            // 
            this.btnCache.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnCache.Image = ((System.Drawing.Image)(resources.GetObject("btnCache.Image")));
            this.btnCache.Label = "Cache Location";
            this.btnCache.Name = "btnCache";
            this.btnCache.ShowImage = true;
            this.btnCache.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonCache_Click);
            // 
            // ribbonZipStatus
            // 
            this.ribbonZipStatus.Items.Add(this.btnActive);
            this.ribbonZipStatus.Label = "Yizip Status";
            this.ribbonZipStatus.Name = "ribbonZipStatus";
            // 
            // btnActive
            // 
            this.btnActive.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnActive.Enabled = false;
            this.btnActive.Image = ((System.Drawing.Image)(resources.GetObject("btnActive.Image")));
            this.btnActive.Label = "Active";
            this.btnActive.Name = "btnActive";
            this.btnActive.ShowImage = true;
            this.btnActive.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonActive_Click);
            // 
            // YizipRibbon
            // 
            this.Name = "YizipRibbon";
            this.RibbonType = "Microsoft.Outlook.Mail.Compose";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.YizipRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.YizipGroup.ResumeLayout(false);
            this.YizipGroup.PerformLayout();
            this.ribbonZipStatus.ResumeLayout(false);
            this.ribbonZipStatus.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup YizipGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnActive;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCache;
        private Microsoft.Office.Tools.Ribbon.RibbonGroup ribbonZipStatus;
    }

    partial class ThisRibbonCollection
    {
        internal YizipRibbon YizipRibbon
        {
            get { return this.GetRibbon<YizipRibbon>(); }
        }
    }
}
