namespace Autocorrect.VSTO
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.AlGrammar = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.correctall = this.Factory.CreateRibbonButton();
            this.correctselected = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.AlGrammar.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "Nertil";
            this.tab1.Label = "Nertil";
            this.tab1.Name = "tab1";
            this.tab1.Visible = false;
            // 
            // AlGrammar
            // 
            this.AlGrammar.Groups.Add(this.group1);
            this.AlGrammar.Label = "Shkruaj Shqip";
            this.AlGrammar.Name = "AlGrammar";
            // 
            // group1
            // 
            this.group1.Items.Add(this.correctall);
            this.group1.Items.Add(this.correctselected);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // correctall
            // 
            this.correctall.Label = "Korrigjo te gjithe dokumentin(Experimental)";
            this.correctall.Name = "correctall";
            this.correctall.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.correctall_Click);
            // 
            // correctselected
            // 
            this.correctselected.Label = "Korrigjo pjesen e zgjedhur(Experimental)";
            this.correctselected.Name = "correctselected";
            this.correctselected.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.correctselected_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.AlGrammar);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.AlGrammar.ResumeLayout(false);
            this.AlGrammar.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        private Microsoft.Office.Tools.Ribbon.RibbonTab AlGrammar;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton correctall;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton correctselected;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
