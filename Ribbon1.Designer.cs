namespace ExcelAddinSCI
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
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
            this.tab2 = this.Factory.CreateRibbonTab();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.RevertToOriginal = this.Factory.CreateRibbonButton();
            this.ConvertToAlphanumeric = this.Factory.CreateRibbonButton();
            this.tab2.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab2
            // 
            this.tab2.Groups.Add(this.group2);
            this.tab2.Label = "tab2";
            this.tab2.Name = "tab2";
            // 
            // group2
            // 
            this.group2.Items.Add(this.ConvertToAlphanumeric);
            this.group2.Items.Add(this.RevertToOriginal);
            this.group2.Label = "group2";
            this.group2.Name = "group2";
            // 
            // RevertToOriginal
            // 
            this.RevertToOriginal.Label = "RevertToOriginal";
            this.RevertToOriginal.Name = "RevertToOriginal";
            this.RevertToOriginal.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RevertToOriginal_Click);
            // 
            // ConvertToAlphanumeric
            // 
            this.ConvertToAlphanumeric.Label = "ConvertToAlphanumeric";
            this.ConvertToAlphanumeric.Name = "ConvertToAlphanumeric";
            this.ConvertToAlphanumeric.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ConvertToAlphanumeric_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab2);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab2.ResumeLayout(false);
            this.tab2.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ConvertToAlphanumeric;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RevertToOriginal;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
