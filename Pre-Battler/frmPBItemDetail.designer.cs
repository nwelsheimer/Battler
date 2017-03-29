namespace Pre_Battler
{
    partial class frmPBItemDetail
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmPBItemDetail));
            this.cmbSKUs = new System.Windows.Forms.ComboBox();
            this.ugrdItemDetail = new Infragistics.Win.UltraWinGrid.UltraGrid();
            this.label1 = new System.Windows.Forms.Label();
            this.btnBack = new System.Windows.Forms.Button();
            this.btnForward = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.ugrdRanks = new Infragistics.Win.UltraWinGrid.UltraGrid();
            this.lblExpand = new System.Windows.Forms.Label();
            this.pnlRank = new System.Windows.Forms.Panel();
            this.btnWOH = new System.Windows.Forms.Button();
            this.btnMin = new System.Windows.Forms.Button();
            this.btnFull = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.lblSite = new System.Windows.Forms.Label();
            this.lblSKUShelf = new System.Windows.Forms.Label();
            this.lblSKUUnits = new System.Windows.Forms.Label();
            this.lblSKU = new System.Windows.Forms.Label();
            this.lblSessionUnits = new System.Windows.Forms.Label();
            this.lblStoreShelf = new System.Windows.Forms.Label();
            this.lblStoreUnits = new System.Windows.Forms.Label();
            this.lblStore = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.cmbWOH = new System.Windows.Forms.ComboBox();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.tOOLSToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.eXPORTTOEXCELToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.rEFRESHAVAILABILITYToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.sELECTCOLUMNSToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.sAVELAYOUTToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.bgExport = new System.ComponentModel.BackgroundWorker();
            this.pnlStatus = new System.Windows.Forms.Panel();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.lblStatus = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.ugrdItemDetail)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.ugrdRanks)).BeginInit();
            this.pnlRank.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.pnlStatus.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // cmbSKUs
            // 
            this.cmbSKUs.FormattingEnabled = true;
            this.cmbSKUs.Location = new System.Drawing.Point(12, 44);
            this.cmbSKUs.Name = "cmbSKUs";
            this.cmbSKUs.Size = new System.Drawing.Size(263, 21);
            this.cmbSKUs.TabIndex = 0;
            this.cmbSKUs.SelectedIndexChanged += new System.EventHandler(this.cmbSKUs_SelectedIndexChanged);
            // 
            // ugrdItemDetail
            // 
            this.ugrdItemDetail.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            appearance1.BackColor = System.Drawing.SystemColors.ControlLight;
            appearance1.BackColor2 = System.Drawing.SystemColors.ControlDark;
            appearance1.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
            appearance1.ForeColorDisabled = System.Drawing.Color.Black;
            this.ugrdItemDetail.DisplayLayout.Appearance = appearance1;
            this.ugrdItemDetail.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ResizeAllColumns;
            this.ugrdItemDetail.DisplayLayout.ColumnChooserEnabled = Infragistics.Win.DefaultableBoolean.True;
            this.ugrdItemDetail.DisplayLayout.GroupByBox.Hidden = true;
            this.ugrdItemDetail.DisplayLayout.InterBandSpacing = 10;
            this.ugrdItemDetail.DisplayLayout.MaxColScrollRegions = 1;
            this.ugrdItemDetail.DisplayLayout.MaxRowScrollRegions = 1;
            this.ugrdItemDetail.DisplayLayout.Override.AllowRowFiltering = Infragistics.Win.DefaultableBoolean.True;
            appearance2.BackColor = System.Drawing.Color.Transparent;
            this.ugrdItemDetail.DisplayLayout.Override.CardAreaAppearance = appearance2;
            appearance3.BackColor = System.Drawing.SystemColors.Control;
            appearance3.BackColor2 = System.Drawing.SystemColors.ControlLightLight;
            appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.VerticalBump;
            this.ugrdItemDetail.DisplayLayout.Override.CellAppearance = appearance3;
            this.ugrdItemDetail.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText;
            appearance4.BackColor = System.Drawing.SystemColors.Control;
            appearance4.BackColor2 = System.Drawing.SystemColors.ControlDark;
            appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.VerticalBump;
            appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
            this.ugrdItemDetail.DisplayLayout.Override.HeaderAppearance = appearance4;
            this.ugrdItemDetail.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
            appearance5.ForeColor = System.Drawing.SystemColors.ActiveCaption;
            this.ugrdItemDetail.DisplayLayout.Override.RowSelectorAppearance = appearance5;
            appearance6.BackColor = System.Drawing.SystemColors.InactiveCaption;
            appearance6.BackColor2 = System.Drawing.SystemColors.ActiveCaption;
            appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.VerticalBump;
            this.ugrdItemDetail.DisplayLayout.Override.SelectedRowAppearance = appearance6;
            this.ugrdItemDetail.DisplayLayout.RowConnectorColor = System.Drawing.SystemColors.ControlDarkDark;
            this.ugrdItemDetail.DisplayLayout.RowConnectorStyle = Infragistics.Win.UltraWinGrid.RowConnectorStyle.Dashed;
            this.ugrdItemDetail.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
            this.ugrdItemDetail.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
            this.ugrdItemDetail.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
            this.ugrdItemDetail.Location = new System.Drawing.Point(1, 154);
            this.ugrdItemDetail.Name = "ugrdItemDetail";
            this.ugrdItemDetail.Size = new System.Drawing.Size(1291, 655);
            this.ugrdItemDetail.TabIndex = 2;
            this.ugrdItemDetail.AfterCellActivate += new System.EventHandler(this.ugrdItemDetail_AfterCellActivate);
            this.ugrdItemDetail.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugrdItemDetail_AfterCellUpdate);
            this.ugrdItemDetail.AfterRowActivate += new System.EventHandler(this.ugrdItemDetail_AfterRowActivate);
            this.ugrdItemDetail.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ugrdItemDetail_KeyDown);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(9, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(94, 17);
            this.label1.TabIndex = 3;
            this.label1.Text = "Select SKU:";
            // 
            // btnBack
            // 
            this.btnBack.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnBack.Location = new System.Drawing.Point(12, 83);
            this.btnBack.Name = "btnBack";
            this.btnBack.Size = new System.Drawing.Size(45, 65);
            this.btnBack.TabIndex = 4;
            this.btnBack.Text = "<-";
            this.btnBack.UseVisualStyleBackColor = true;
            this.btnBack.Click += new System.EventHandler(this.btnBack_Click);
            // 
            // btnForward
            // 
            this.btnForward.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnForward.Location = new System.Drawing.Point(230, 83);
            this.btnForward.Name = "btnForward";
            this.btnForward.Size = new System.Drawing.Size(45, 65);
            this.btnForward.TabIndex = 5;
            this.btnForward.Text = "->";
            this.btnForward.UseVisualStyleBackColor = true;
            this.btnForward.Click += new System.EventHandler(this.btnForward_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(63, 108);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(160, 17);
            this.label2.TabIndex = 6;
            this.label2.Text = "Advance to next SKU";
            // 
            // ugrdRanks
            // 
            this.ugrdRanks.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            appearance7.BackColor = System.Drawing.SystemColors.ControlLight;
            appearance7.BackColor2 = System.Drawing.SystemColors.ControlDark;
            appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
            appearance7.ForeColorDisabled = System.Drawing.Color.Black;
            this.ugrdRanks.DisplayLayout.Appearance = appearance7;
            this.ugrdRanks.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ResizeAllColumns;
            this.ugrdRanks.DisplayLayout.InterBandSpacing = 10;
            this.ugrdRanks.DisplayLayout.MaxColScrollRegions = 1;
            this.ugrdRanks.DisplayLayout.MaxRowScrollRegions = 1;
            appearance8.BackColor = System.Drawing.Color.Transparent;
            this.ugrdRanks.DisplayLayout.Override.CardAreaAppearance = appearance8;
            appearance9.BackColor = System.Drawing.SystemColors.Control;
            appearance9.BackColor2 = System.Drawing.SystemColors.ControlLightLight;
            appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.VerticalBump;
            this.ugrdRanks.DisplayLayout.Override.CellAppearance = appearance9;
            this.ugrdRanks.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText;
            appearance10.BackColor = System.Drawing.SystemColors.Control;
            appearance10.BackColor2 = System.Drawing.SystemColors.ControlDark;
            appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.VerticalBump;
            appearance10.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
            this.ugrdRanks.DisplayLayout.Override.HeaderAppearance = appearance10;
            this.ugrdRanks.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
            appearance11.ForeColor = System.Drawing.SystemColors.ActiveCaption;
            this.ugrdRanks.DisplayLayout.Override.RowSelectorAppearance = appearance11;
            appearance12.BackColor = System.Drawing.SystemColors.InactiveCaption;
            appearance12.BackColor2 = System.Drawing.SystemColors.ActiveCaption;
            appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.VerticalBump;
            this.ugrdRanks.DisplayLayout.Override.SelectedRowAppearance = appearance12;
            this.ugrdRanks.DisplayLayout.RowConnectorColor = System.Drawing.SystemColors.ControlDarkDark;
            this.ugrdRanks.DisplayLayout.RowConnectorStyle = Infragistics.Win.UltraWinGrid.RowConnectorStyle.Dashed;
            this.ugrdRanks.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
            this.ugrdRanks.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
            this.ugrdRanks.Location = new System.Drawing.Point(3, 3);
            this.ugrdRanks.Name = "ugrdRanks";
            this.ugrdRanks.Size = new System.Drawing.Size(207, 116);
            this.ugrdRanks.TabIndex = 7;
            this.ugrdRanks.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ugrdRanks_KeyDown);
            // 
            // lblExpand
            // 
            this.lblExpand.AutoSize = true;
            this.lblExpand.ForeColor = System.Drawing.Color.Blue;
            this.lblExpand.Location = new System.Drawing.Point(310, 8);
            this.lblExpand.Name = "lblExpand";
            this.lblExpand.Size = new System.Drawing.Size(203, 13);
            this.lblExpand.TabIndex = 8;
            this.lblExpand.Text = "Enter quantities by rank. Press to expand.";
            this.lblExpand.Click += new System.EventHandler(this.lblExpand_Click);
            // 
            // pnlRank
            // 
            this.pnlRank.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnlRank.Controls.Add(this.ugrdRanks);
            this.pnlRank.Location = new System.Drawing.Point(309, 24);
            this.pnlRank.Name = "pnlRank";
            this.pnlRank.Size = new System.Drawing.Size(215, 124);
            this.pnlRank.TabIndex = 9;
            // 
            // btnWOH
            // 
            this.btnWOH.Location = new System.Drawing.Point(50, 43);
            this.btnWOH.Name = "btnWOH";
            this.btnWOH.Size = new System.Drawing.Size(93, 36);
            this.btnWOH.TabIndex = 10;
            this.btnWOH.Text = "Distribute WOH";
            this.btnWOH.UseVisualStyleBackColor = true;
            this.btnWOH.Click += new System.EventHandler(this.btnWOH_Click);
            // 
            // btnMin
            // 
            this.btnMin.Location = new System.Drawing.Point(631, 112);
            this.btnMin.Name = "btnMin";
            this.btnMin.Size = new System.Drawing.Size(95, 36);
            this.btnMin.TabIndex = 11;
            this.btnMin.Text = "Distribute Min";
            this.btnMin.UseVisualStyleBackColor = true;
            this.btnMin.Click += new System.EventHandler(this.btnMin_Click);
            // 
            // btnFull
            // 
            this.btnFull.Location = new System.Drawing.Point(530, 112);
            this.btnFull.Name = "btnFull";
            this.btnFull.Size = new System.Drawing.Size(95, 36);
            this.btnFull.TabIndex = 12;
            this.btnFull.Text = "Distribute Full";
            this.btnFull.UseVisualStyleBackColor = true;
            this.btnFull.Click += new System.EventHandler(this.btnFull_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.lblSite);
            this.groupBox1.Controls.Add(this.lblSKUShelf);
            this.groupBox1.Controls.Add(this.lblSKUUnits);
            this.groupBox1.Controls.Add(this.lblSKU);
            this.groupBox1.Controls.Add(this.lblSessionUnits);
            this.groupBox1.Controls.Add(this.lblStoreShelf);
            this.groupBox1.Controls.Add(this.lblStoreUnits);
            this.groupBox1.Controls.Add(this.lblStore);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(869, 21);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(413, 127);
            this.groupBox1.TabIndex = 13;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Summary Information";
            // 
            // lblSite
            // 
            this.lblSite.AutoSize = true;
            this.lblSite.Location = new System.Drawing.Point(326, 100);
            this.lblSite.Name = "lblSite";
            this.lblSite.Size = new System.Drawing.Size(81, 13);
            this.lblSite.TabIndex = 20;
            this.lblSite.Text = "Ship Site: AL";
            // 
            // lblSKUShelf
            // 
            this.lblSKUShelf.AutoSize = true;
            this.lblSKUShelf.Location = new System.Drawing.Point(225, 59);
            this.lblSKUShelf.Name = "lblSKUShelf";
            this.lblSKUShelf.Size = new System.Drawing.Size(91, 13);
            this.lblSKUShelf.TabIndex = 19;
            this.lblSKUShelf.Text = "Total Shelf: 18";
            // 
            // lblSKUUnits
            // 
            this.lblSKUUnits.AutoSize = true;
            this.lblSKUUnits.Location = new System.Drawing.Point(225, 40);
            this.lblSKUUnits.Name = "lblSKUUnits";
            this.lblSKUUnits.Size = new System.Drawing.Size(105, 13);
            this.lblSKUUnits.TabIndex = 18;
            this.lblSKUUnits.Text = "Total Units: 2100";
            // 
            // lblSKU
            // 
            this.lblSKU.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblSKU.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSKU.Location = new System.Drawing.Point(214, 20);
            this.lblSKU.Name = "lblSKU";
            this.lblSKU.Size = new System.Drawing.Size(193, 62);
            this.lblSKU.TabIndex = 17;
            this.lblSKU.Text = "AWESOME COLOR 2.5 QT";
            this.lblSKU.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // lblSessionUnits
            // 
            this.lblSessionUnits.AutoSize = true;
            this.lblSessionUnits.Location = new System.Drawing.Point(15, 100);
            this.lblSessionUnits.Name = "lblSessionUnits";
            this.lblSessionUnits.Size = new System.Drawing.Size(144, 13);
            this.lblSessionUnits.TabIndex = 16;
            this.lblSessionUnits.Text = "Total Requested: 86510";
            // 
            // lblStoreShelf
            // 
            this.lblStoreShelf.AutoSize = true;
            this.lblStoreShelf.Location = new System.Drawing.Point(15, 59);
            this.lblStoreShelf.Name = "lblStoreShelf";
            this.lblStoreShelf.Size = new System.Drawing.Size(91, 13);
            this.lblStoreShelf.TabIndex = 2;
            this.lblStoreShelf.Text = "Total Shelf: 18";
            // 
            // lblStoreUnits
            // 
            this.lblStoreUnits.AutoSize = true;
            this.lblStoreUnits.Location = new System.Drawing.Point(15, 40);
            this.lblStoreUnits.Name = "lblStoreUnits";
            this.lblStoreUnits.Size = new System.Drawing.Size(105, 13);
            this.lblStoreUnits.TabIndex = 1;
            this.lblStoreUnits.Text = "Total Units: 2100";
            // 
            // lblStore
            // 
            this.lblStore.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblStore.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblStore.Location = new System.Drawing.Point(6, 20);
            this.lblStore.Name = "lblStore";
            this.lblStore.Size = new System.Drawing.Size(193, 62);
            this.lblStore.TabIndex = 0;
            this.lblStore.Text = "SAMPLE STORE 1122";
            this.lblStore.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.cmbWOH);
            this.groupBox2.Controls.Add(this.btnWOH);
            this.groupBox2.Location = new System.Drawing.Point(530, 21);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(195, 85);
            this.groupBox2.TabIndex = 14;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "WOH Calculation Method";
            // 
            // cmbWOH
            // 
            this.cmbWOH.FormattingEnabled = true;
            this.cmbWOH.Items.AddRange(new object[] {
            "WOH last week",
            "WOH this week",
            "WOH 2 weeks ago",
            "WOH 2 week average"});
            this.cmbWOH.Location = new System.Drawing.Point(6, 16);
            this.cmbWOH.Name = "cmbWOH";
            this.cmbWOH.Size = new System.Drawing.Size(183, 21);
            this.cmbWOH.TabIndex = 11;
            this.cmbWOH.Text = "WOH last week";
            this.cmbWOH.SelectedIndexChanged += new System.EventHandler(this.cmbWOH_SelectedIndexChanged);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tOOLSToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1293, 24);
            this.menuStrip1.TabIndex = 15;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // tOOLSToolStripMenuItem
            // 
            this.tOOLSToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.eXPORTTOEXCELToolStripMenuItem,
            this.rEFRESHAVAILABILITYToolStripMenuItem,
            this.sELECTCOLUMNSToolStripMenuItem,
            this.sAVELAYOUTToolStripMenuItem});
            this.tOOLSToolStripMenuItem.Name = "tOOLSToolStripMenuItem";
            this.tOOLSToolStripMenuItem.Size = new System.Drawing.Size(55, 20);
            this.tOOLSToolStripMenuItem.Text = "TOOLS";
            // 
            // eXPORTTOEXCELToolStripMenuItem
            // 
            this.eXPORTTOEXCELToolStripMenuItem.Name = "eXPORTTOEXCELToolStripMenuItem";
            this.eXPORTTOEXCELToolStripMenuItem.Size = new System.Drawing.Size(195, 22);
            this.eXPORTTOEXCELToolStripMenuItem.Text = "EXPORT TO EXCEL";
            this.eXPORTTOEXCELToolStripMenuItem.Click += new System.EventHandler(this.eXPORTTOEXCELToolStripMenuItem_Click);
            // 
            // rEFRESHAVAILABILITYToolStripMenuItem
            // 
            this.rEFRESHAVAILABILITYToolStripMenuItem.Name = "rEFRESHAVAILABILITYToolStripMenuItem";
            this.rEFRESHAVAILABILITYToolStripMenuItem.Size = new System.Drawing.Size(195, 22);
            this.rEFRESHAVAILABILITYToolStripMenuItem.Text = "REFRESH AVAILABILITY";
            this.rEFRESHAVAILABILITYToolStripMenuItem.Click += new System.EventHandler(this.rEFRESHAVAILABILITYToolStripMenuItem_Click);
            // 
            // sELECTCOLUMNSToolStripMenuItem
            // 
            this.sELECTCOLUMNSToolStripMenuItem.Name = "sELECTCOLUMNSToolStripMenuItem";
            this.sELECTCOLUMNSToolStripMenuItem.Size = new System.Drawing.Size(195, 22);
            this.sELECTCOLUMNSToolStripMenuItem.Text = "SELECT COLUMNS";
            this.sELECTCOLUMNSToolStripMenuItem.Click += new System.EventHandler(this.sELECTCOLUMNSToolStripMenuItem_Click);
            // 
            // sAVELAYOUTToolStripMenuItem
            // 
            this.sAVELAYOUTToolStripMenuItem.Name = "sAVELAYOUTToolStripMenuItem";
            this.sAVELAYOUTToolStripMenuItem.Size = new System.Drawing.Size(195, 22);
            this.sAVELAYOUTToolStripMenuItem.Text = "SAVE LAYOUT";
            this.sAVELAYOUTToolStripMenuItem.Click += new System.EventHandler(this.sAVELAYOUTToolStripMenuItem_Click);
            // 
            // bgExport
            // 
            this.bgExport.DoWork += new System.ComponentModel.DoWorkEventHandler(this.bgExport_DoWork);
            this.bgExport.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.bgExport_RunWorkerCompleted);
            // 
            // pnlStatus
            // 
            this.pnlStatus.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnlStatus.Controls.Add(this.lblStatus);
            this.pnlStatus.Controls.Add(this.pictureBox1);
            this.pnlStatus.Location = new System.Drawing.Point(732, 28);
            this.pnlStatus.Name = "pnlStatus";
            this.pnlStatus.Size = new System.Drawing.Size(119, 119);
            this.pnlStatus.TabIndex = 16;
            this.pnlStatus.Visible = false;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::Pre_Battler.Properties.Resources._58c7ec502dc4e637388739;
            this.pictureBox1.Location = new System.Drawing.Point(3, 36);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(112, 80);
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // lblStatus
            // 
            this.lblStatus.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblStatus.Location = new System.Drawing.Point(3, 1);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(116, 33);
            this.lblStatus.TabIndex = 1;
            this.lblStatus.Text = "Background export in progress";
            this.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // frmPBItemDetail
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1293, 808);
            this.Controls.Add(this.pnlStatus);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnFull);
            this.Controls.Add(this.btnMin);
            this.Controls.Add(this.pnlRank);
            this.Controls.Add(this.lblExpand);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnForward);
            this.Controls.Add(this.btnBack);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.ugrdItemDetail);
            this.Controls.Add(this.cmbSKUs);
            this.Controls.Add(this.menuStrip1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmPBItemDetail";
            this.Text = "Distribute SKUs";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            ((System.ComponentModel.ISupportInitialize)(this.ugrdItemDetail)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.ugrdRanks)).EndInit();
            this.pnlRank.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.pnlStatus.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox cmbSKUs;
        private Infragistics.Win.UltraWinGrid.UltraGrid ugrdItemDetail;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnBack;
        private System.Windows.Forms.Button btnForward;
        private System.Windows.Forms.Label label2;
        private Infragistics.Win.UltraWinGrid.UltraGrid ugrdRanks;
        private System.Windows.Forms.Label lblExpand;
        private System.Windows.Forms.Panel pnlRank;
        private System.Windows.Forms.Button btnWOH;
        private System.Windows.Forms.Button btnMin;
        private System.Windows.Forms.Button btnFull;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label lblStoreUnits;
        private System.Windows.Forms.Label lblStore;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.ComboBox cmbWOH;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem tOOLSToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem sAVELAYOUTToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem sELECTCOLUMNSToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem eXPORTTOEXCELToolStripMenuItem;
        private System.Windows.Forms.Label lblSessionUnits;
        private System.Windows.Forms.Label lblStoreShelf;
        private System.Windows.Forms.Label lblSKU;
        private System.Windows.Forms.ToolStripMenuItem rEFRESHAVAILABILITYToolStripMenuItem;
        private System.Windows.Forms.Label lblSKUShelf;
        private System.Windows.Forms.Label lblSKUUnits;
        private System.Windows.Forms.Label lblSite;
        private System.ComponentModel.BackgroundWorker bgExport;
        private System.Windows.Forms.Panel pnlStatus;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label lblStatus;
    }
}