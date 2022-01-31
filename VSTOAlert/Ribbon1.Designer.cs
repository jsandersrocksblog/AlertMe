namespace VSTOAlert
{
    partial class AlertRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public AlertRibbon()
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
            this.ToAddresses = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.gallery1 = this.Factory.CreateRibbonGallery();
            this.btnClearAll = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.editBox1 = this.Factory.CreateRibbonEditBox();
            this.btnAdd = this.Factory.CreateRibbonButton();
            this.ToAddresses.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // ToAddresses
            // 
            this.ToAddresses.Groups.Add(this.group1);
            this.ToAddresses.Groups.Add(this.group2);
            this.ToAddresses.Label = "Alert Me";
            this.ToAddresses.Name = "ToAddresses";
            // 
            // group1
            // 
            this.group1.Items.Add(this.gallery1);
            this.group1.Items.Add(this.btnClearAll);
            this.group1.Label = "Click on Item to remove";
            this.group1.Name = "group1";
            // 
            // gallery1
            // 
            this.gallery1.ColumnCount = 4;
            this.gallery1.Label = "Remove Item";
            this.gallery1.Name = "gallery1";
            this.gallery1.RowCount = 4;
            this.gallery1.ShowItemSelection = true;
            this.gallery1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.gallery1_Click);
            // 
            // btnClearAll
            // 
            this.btnClearAll.Label = "Remove All";
            this.btnClearAll.Name = "btnClearAll";
            this.btnClearAll.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.clearAllClick);
            // 
            // group2
            // 
            this.group2.Items.Add(this.editBox1);
            this.group2.Items.Add(this.btnAdd);
            this.group2.Label = "Add Item";
            this.group2.Name = "group2";
            // 
            // editBox1
            // 
            this.editBox1.Label = "editBox1";
            this.editBox1.Name = "editBox1";
            this.editBox1.ShowLabel = false;
            this.editBox1.Text = null;
            // 
            // btnAdd
            // 
            this.btnAdd.Label = "Add";
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.addClick);
            // 
            // AlertRibbon
            // 
            this.Name = "AlertRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.ToAddresses);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.ToAddresses.ResumeLayout(false);
            this.ToAddresses.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab ToAddresses;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery gallery1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAdd;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnClearAll;
    }

    partial class ThisRibbonCollection
    {
        internal AlertRibbon Ribbon1
        {
            get { return this.GetRibbon<AlertRibbon>(); }
        }
    }
}
