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
            this.ActionGroup = this.Factory.CreateRibbonGroup();
            this.autocorrectToggle = this.Factory.CreateRibbonToggleButton();
            this.licensing = this.Factory.CreateRibbonGroup();
            this.license = this.Factory.CreateRibbonButton();
            this.LicenseDetails = this.Factory.CreateRibbonGroup();
            this.expirationDateLable = this.Factory.CreateRibbonLabel();
            this.expirationDateValueLabel = this.Factory.CreateRibbonLabel();
            this.hasExpired = this.Factory.CreateRibbonLabel();
            this.tab1.SuspendLayout();
            this.AlGrammar.SuspendLayout();
            this.group1.SuspendLayout();
            this.ActionGroup.SuspendLayout();
            this.licensing.SuspendLayout();
            this.LicenseDetails.SuspendLayout();
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
            this.AlGrammar.Groups.Add(this.ActionGroup);
            this.AlGrammar.Groups.Add(this.licensing);
            this.AlGrammar.Groups.Add(this.LicenseDetails);
            this.AlGrammar.Label = "TekstSakte";
            this.AlGrammar.Name = "AlGrammar";
            // 
            // group1
            // 
            this.group1.Items.Add(this.correctall);
            this.group1.Items.Add(this.correctselected);
            this.group1.Label = "Korrektim";
            this.group1.Name = "group1";
            // 
            // correctall
            // 
            this.correctall.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.correctall.Image = global::Autocorrect.VSTO.Properties.Resources.grammarfix;
            this.correctall.Label = "Korrigjo te gjithe";
            this.correctall.Name = "correctall";
            this.correctall.ShowImage = true;
            this.correctall.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.correctall_Click);
            // 
            // correctselected
            // 
            this.correctselected.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.correctselected.Image = global::Autocorrect.VSTO.Properties.Resources.grammarfix;
            this.correctselected.Label = "Korrigjo zgjedhjen";
            this.correctselected.Name = "correctselected";
            this.correctselected.ShowImage = true;
            this.correctselected.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.correctselected_Click);
            // 
            // ActionGroup
            // 
            this.ActionGroup.Items.Add(this.autocorrectToggle);
            this.ActionGroup.Name = "ActionGroup";
            // 
            // autocorrectToggle
            // 
            this.autocorrectToggle.Checked = true;
            this.autocorrectToggle.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.autocorrectToggle.Image = global::Autocorrect.VSTO.Properties.Resources.autocorrect;
            this.autocorrectToggle.Label = "Korrigjo Automatikish";
            this.autocorrectToggle.Name = "autocorrectToggle";
            this.autocorrectToggle.ShowImage = true;
            this.autocorrectToggle.SuperTip = "Kur eshte aktivizuar korrigjon tekstin nderkohe qe po shkruhet";
            this.autocorrectToggle.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.autocorrectToggle_Click);
            // 
            // licensing
            // 
            this.licensing.Items.Add(this.license);
            this.licensing.Label = "Licensim";
            this.licensing.Name = "licensing";
            // 
            // license
            // 
            this.license.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.license.Image = global::Autocorrect.VSTO.Properties.Resources.icon_license_keys;
            this.license.Label = "Rregjistohu";
            this.license.Name = "license";
            this.license.ShowImage = true;
            this.license.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.license_Click);
            // 
            // LicenseDetails
            // 
            this.LicenseDetails.Items.Add(this.expirationDateLable);
            this.LicenseDetails.Items.Add(this.expirationDateValueLabel);
            this.LicenseDetails.Items.Add(this.hasExpired);
            this.LicenseDetails.Label = "Licensa";
            this.LicenseDetails.Name = "LicenseDetails";
            // 
            // expirationDateLable
            // 
            this.expirationDateLable.Label = "Data Skadimit";
            this.expirationDateLable.Name = "expirationDateLable";
            // 
            // expirationDateValueLabel
            // 
            this.expirationDateValueLabel.Label = "--/--/--";
            this.expirationDateValueLabel.Name = "expirationDateValueLabel";
            // 
            // hasExpired
            // 
            this.hasExpired.Label = "Valid";
            this.hasExpired.Name = "hasExpired";
            this.hasExpired.Visible = false;
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
            this.ActionGroup.ResumeLayout(false);
            this.ActionGroup.PerformLayout();
            this.licensing.ResumeLayout(false);
            this.licensing.PerformLayout();
            this.LicenseDetails.ResumeLayout(false);
            this.LicenseDetails.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        private Microsoft.Office.Tools.Ribbon.RibbonTab AlGrammar;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton correctall;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton correctselected;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup licensing;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton license;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup LicenseDetails;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel expirationDateLable;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel expirationDateValueLabel;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel hasExpired;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup ActionGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton autocorrectToggle;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
